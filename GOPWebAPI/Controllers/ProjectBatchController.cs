using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using System.Data.SqlTypes;
using System.Net.Http.Headers;
using GOPWebAPI.Helpers;
using Aspose.Cells;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/projectbatch")]
    public class ProjectBatchController : ApiController
    {
        #region Create Project Batch
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateProjectBatch([FromBody]ProjectBatch projectBatch)
        {
            try
            {
                string CustomerInputFileExtension = string.Empty;

                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Batch Data");

                if (!AccessControl.CanUserAccessPage(projectBatch.UserID, "Create Project Batch"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Get Customer Input File Extension from Temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                if (File.Exists(dirTemp + projectBatch.CustomerInputFileName))
                    CustomerInputFileExtension = Path.GetExtension(dirTemp + projectBatch.CustomerInputFileName);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectBatch";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", projectBatch.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectBatch.ProjectCode);
                    cmd.Parameters.AddWithValue("@InputCount", projectBatch.InputCount);
                    cmd.Parameters.AddWithValue("@InputCountType", projectBatch.InputCountType);
                    cmd.Parameters.AddWithValue("@ReceivedDate", projectBatch.ReceivedDate);
                    cmd.Parameters.AddWithValue("@ReceivedFormat", projectBatch.ReceivedFormat);
                    cmd.Parameters.AddWithValue("@OutputFormat", projectBatch.OutputFormat);
                    cmd.Parameters.AddWithValue("@PlannedStartDate", projectBatch.PlannedStartDate);
                    cmd.Parameters.AddWithValue("@PlannedDeliveryDate", projectBatch.PlannedDeliveryDate);
                    cmd.Parameters.AddWithValue("@Remarks", projectBatch.Remarks);
                    cmd.Parameters.AddWithValue("@CustomerInputFileExtension", CustomerInputFileExtension);
                    cmd.Parameters.AddWithValue("@UserID", projectBatch.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);
                    #endregion

                    //Calling sp to create project batch
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "created")
                    {
                        string BatchNo = arrResult[1];

                        #region Moving Customer Input File
                        if (File.Exists(dirTemp + projectBatch.CustomerInputFileName))
                        {
                            string FinalCIFName = projectBatch.CustomerCode + "_" + projectBatch.ProjectCode + "_" + BatchNo + "_CustomerInputFile" + CustomerInputFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));
                            FileOperations.MoveFile(dirTemp, projectBatch.CustomerInputFileName, dirUploads, FinalCIFName);
                        }
                        #endregion

                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, BatchNo);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Project Batch List By Project Id
        [HttpGet]
        [Route("projectbatchlist/{id}/{UserID}/{status?}")]
        public IHttpActionResult ReadProjectBatchListByProjectId(long id,string UserID, string Status = "O") //O-Ongoing, D-Delivered
        {
            try
            {
                bool canUserDeliverProjectBatch = false;
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Project Batch List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                if (AccessControl.CanUserAccessPage(UserID, "Deliver Project Batch"))
                    canUserDeliverProjectBatch = true;

                //Create a list to hold the list of Project Batches
                List<ProjectBatch> ProjectBatchList = new List<ProjectBatch>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectBatchList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", id);
                    cmd.Parameters.AddWithValue("@Status", Status);

                    //Calling sp to get list of project batches
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        ProjectBatch projectBatch = new ProjectBatch();

                        //Assign values to model object
                        projectBatch.SlNo = SlNo;
                        projectBatch.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        projectBatch.ProjectBatchID = Convert.ToInt64(sqlReader["ProjectBatchID"]);
                        projectBatch.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectBatch.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectBatch.BatchNo = sqlReader["BatchNo"].ToString();
                        projectBatch.Scope = sqlReader["Scope"].ToString();
                        projectBatch.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        projectBatch.PlannedDeliveryDate = Convert.ToDateTime(sqlReader["PlannedDeliveryDate"]);
                        projectBatch.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        projectBatch.InputCountType = sqlReader["InputCountType"].ToString();
                        projectBatch.Status = sqlReader["Status"].ToString();
                        if (Status == "D")
                        {
                            projectBatch.DeliveredDate = Convert.ToDateTime(sqlReader["DeliveredDate"]);
                            projectBatch.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        }
                        projectBatch.canUserDeliverProjectBatch = canUserDeliverProjectBatch;
                        projectBatch.IsPostProjectBatchDetailsExist = Convert.ToBoolean(sqlReader["IsPostProjectBatchDetailsExist"]);

                        //Add object to list
                        ProjectBatchList.Add(projectBatch);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectBatchList);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Project Batch By Project Batch Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadProjectBatchByProjectBatchId(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Project Batch"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a project batch instance
                ProjectBatch projectBatch = new ProjectBatch();
                System.Data.Common.DbDataReader sqlReader;
                                
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectBatch";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@ProjectBatchID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get project batch details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        //Assign data to project batch object properties
                        projectBatch.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        projectBatch.ProjectBatchID = Convert.ToInt64(sqlReader["ProjectBatchID"]);
                        projectBatch.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectBatch.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectBatch.Scope = sqlReader["Scope"].ToString();
                        projectBatch.BatchNo = sqlReader["BatchNo"].ToString();
                        projectBatch.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        projectBatch.InputCountType = sqlReader["InputCountType"].ToString();
                        projectBatch.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        projectBatch.ReceivedFormat = sqlReader["ReceivedFormat"].ToString();
                        projectBatch.OutputFormat = sqlReader["OutputFormat"].ToString();
                        projectBatch.PlannedStartDate = Convert.ToDateTime(sqlReader["PlannedStartDate"]);
                        projectBatch.PlannedDeliveryDate = Convert.ToDateTime(sqlReader["PlannedDeliveryDate"]);
                        projectBatch.Remarks = sqlReader["Remarks"].ToString();
                        projectBatch.CustomerInputFileName = sqlReader["CustomerInputFileName"].ToString();
                        projectBatch.DeliveredDate = sqlReader["DeliveredOn"] == DBNull.Value ? null : (DateTime?)sqlReader["DeliveredOn"];
                        projectBatch.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        projectBatch.Status = sqlReader["ProjectBatchStatus"].ToString();
                        conn.Close();

                        //return project batch to the request
                        return Ok(projectBatch);
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }
                    
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Update Project Batch
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateProjectBatch(int id, [FromBody]ProjectBatch projectBatch)
        {
            try
            {
                string CustomerInputFileExtension = string.Empty;

                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Project Batch Id
                if (id != projectBatch.ProjectBatchID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Batch Id");

                if (!AccessControl.CanUserAccessPage(projectBatch.UserID, "Edit Project Batch"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Get the customer input file extension from temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                //Customer Input File
                if (File.Exists(dirTemp + projectBatch.CustomerInputFileName))
                    CustomerInputFileExtension = Path.GetExtension(dirTemp + projectBatch.CustomerInputFileName);
                else if (!string.IsNullOrEmpty(projectBatch.CustomerInputFileName))
                    CustomerInputFileExtension = projectBatch.CustomerInputFileName.Substring(projectBatch.CustomerInputFileName.LastIndexOf('.'), projectBatch.CustomerInputFileName.Length - projectBatch.CustomerInputFileName.LastIndexOf('.'));
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectBatch";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@ProjectBatchID", projectBatch.ProjectBatchID);
                    cmd.Parameters.AddWithValue("@CustomerCode", projectBatch.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectBatch.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", projectBatch.BatchNo);
                    cmd.Parameters.AddWithValue("@InputCount", projectBatch.InputCount);
                    cmd.Parameters.AddWithValue("@InputCountType", projectBatch.InputCountType);
                    cmd.Parameters.AddWithValue("@ReceivedDate", projectBatch.ReceivedDate);
                    cmd.Parameters.AddWithValue("@ReceivedFormat", projectBatch.ReceivedFormat);
                    cmd.Parameters.AddWithValue("@OutputFormat", projectBatch.OutputFormat);
                    cmd.Parameters.AddWithValue("@PlannedStartDate", projectBatch.PlannedStartDate);
                    cmd.Parameters.AddWithValue("@PlannedDeliveryDate", projectBatch.PlannedDeliveryDate);
                    cmd.Parameters.AddWithValue("@Remarks", projectBatch.Remarks);
                    cmd.Parameters.AddWithValue("@CustomerInputFileExtension", CustomerInputFileExtension);
                    cmd.Parameters.AddWithValue("@UserID", projectBatch.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);
                    #endregion

                    //Calling sp to update project batch
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();
                    if (Result.Trim().ToLower() == "updated")
                    {
                        #region Move Customer Input File
                        DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));
                        if (File.Exists(dirTemp + projectBatch.CustomerInputFileName))
                        {
                            string FinalCIFName = projectBatch.CustomerCode + "_" + projectBatch.ProjectCode + "_" + projectBatch.BatchNo + "_CustomerInputFile" + CustomerInputFileExtension;
                            FileOperations.MoveFile(dirTemp, projectBatch.CustomerInputFileName, dirCIFUploads, FinalCIFName);
                        }
                        //since customer input file is mandatory no else clause is required
                        #endregion

                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }
        #endregion
        
        #region Delete Project Batch
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteProjectBatch(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Project Batch"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectBatch";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@ProjectBatchID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Project Batch
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');

                    if (arrResult[0].ToLower() == "deleted")
                    {
                        string CustomerCode = arrResult[1];
                        string ProjectCode = arrResult[2];
                        string BatchNo = arrResult[3];

                        #region Delete Customer Input Batch File
                        DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));
                        foreach (FileInfo file in dirCIFUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_" + BatchNo.ToUpper() +"_CUSTOMERINPUTFILE"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Download Customer Input Batch File
        [HttpGet]
        [Route("downloadcustomerinputbatchfile")]
        public HttpResponseMessage DownloadCustomerInputBatchFile(string FileName)
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path
                string filePath = HttpContext.Current.Server.MapPath("~/temp/") + FileName;

                if (!File.Exists(filePath))
                {
                    filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/") + FileName;

                    //Check whether File exists
                    if (!File.Exists(filePath))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");
                }

                //Read the File into a Byte Array
                byte[] bytes = File.ReadAllBytes(filePath);

                //Set the Response Content
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                //Set the File Content Type
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileName));

                return response;
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Change Project Batch No.
        [HttpPost]
        [Route("ChangeProjectBatchNo/{CustomerCode}/{ProjectCode}/{BatchNo}/{ChangeToBatchNo}/{UserID}")]
        public HttpResponseMessage ChangeProjectBatchNo(string CustomerCode, string ProjectCode, string BatchNo, string ChangeToBatchNo, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Change Project Batch No"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spChangeProjectBatchNo";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ChangeToBatchNo", ChangeToBatchNo);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to update project batch no.
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "updated")
                    {
                        string ChangedCustomerInputFileName = arrResult[1];

                        #region Renaming Project Batch Customer Input File
                        if (!string.IsNullOrEmpty(ChangedCustomerInputFileName))
                        {
                            DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));
                            foreach (FileInfo file in dirCIFUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_" + BatchNo.ToUpper() + "_CUSTOMERINPUTFILE"))
                                {
                                    file.MoveTo(Path.Combine(dirCIFUploads.FullName, ChangedCustomerInputFileName));
                                    break;
                                }
                            }
                        }
                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Export Project Batch List To Excel
        [HttpGet]
        [Route("ExportProjectBatchListToExcel/{id}/{status?}")]
        public HttpResponseMessage ExportProjectBatchListToExcel(long id, string Status = "O")
        {
            try
            {
                string FileName = "ProjectBatchList.xlsx";

                //Create a list to hold the list of Project Batches
                List<ProjectBatch> ProjectBatchList = new List<ProjectBatch>();
                System.Data.Common.DbDataReader sqlReader;

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                #region Setting up the workbook
                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;
                #endregion

                #region Setting Styles
                Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();

                styleHeader.IsTextWrapped = true;
                styleHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleHeader.VerticalAlignment = TextAlignmentType.Center;
                styleHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                styleHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                styleHeader.Pattern = BackgroundType.VerticalStripe;
                styleHeader.Font.Color = System.Drawing.Color.Black;
                styleHeader.Font.IsBold = true;
                styleHeader.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleHeader.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleHeader.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleHeader.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                styleCenterAlignData.HorizontalAlignment = TextAlignmentType.Center;
                styleCenterAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                styleLeftAlignData.HorizontalAlignment = TextAlignmentType.Left;
                styleLeftAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleLeftAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleLeftAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleLeftAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleLeftAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleLeftAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleLeftAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleLeftAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectBatchList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", id);
                    cmd.Parameters.AddWithValue("@Status", Status);

                    //Calling sp to get list of project batches
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        #region Adding data to List
                        while (sqlReader.Read())
                        {
                            //Create model object
                            ProjectBatch projectBatch = new ProjectBatch();

                            //Assign values to model object
                            projectBatch.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                            projectBatch.ProjectBatchID = Convert.ToInt64(sqlReader["ProjectBatchID"]);
                            projectBatch.CustomerCode = sqlReader["CustomerCode"].ToString();
                            projectBatch.ProjectCode = sqlReader["ProjectCode"].ToString();
                            projectBatch.BatchNo = sqlReader["BatchNo"].ToString();
                            projectBatch.Scope = sqlReader["Scope"].ToString();
                            projectBatch.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                            projectBatch.PlannedDeliveryDate = Convert.ToDateTime(sqlReader["PlannedDeliveryDate"]);
                            projectBatch.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                            projectBatch.InputCountType = sqlReader["InputCountType"].ToString();
                            if (Status == "D")
                            {
                                projectBatch.DeliveredDate = Convert.ToDateTime(sqlReader["DeliveredDate"]);
                                projectBatch.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                            }
                            projectBatch.Status = sqlReader["Status"].ToString();

                            //Add object to list
                            ProjectBatchList.Add(projectBatch);
                        }
                        #endregion
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Customer Code");
                        ws.Cells[0, 2].PutValue("Project Code");
                        ws.Cells[0, 3].PutValue("Batch No.");
                        ws.Cells[0, 4].PutValue("Scope");
                        ws.Cells[0, 5].PutValue("Received Date");
                        ws.Cells[0, 6].PutValue("Planned Delivery Date");
                        ws.Cells[0, 7].PutValue("Input Count");
                        ws.Cells[0, 8].PutValue("Input Count Type");
                        ws.Cells[0, 9].PutValue("Status");
                        if (Status == "D")
                        {
                            ws.Cells[0, 10].PutValue("Delivered Date");
                            ws.Cells[0, 11].PutValue("Delivered Count");
                        }

                        if (Status == "O")
                        {
                            for (int c = 0; c <= 9; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                        }
                        else
                        {
                            for (int c = 0; c <= 11; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                        }
                        #endregion

                        #region Writing row data
                        foreach (ProjectBatch projectBatch in ProjectBatchList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(projectBatch.CustomerCode);
                            ws.Cells[row, 2].PutValue(projectBatch.ProjectCode);
                            ws.Cells[row, 3].PutValue(projectBatch.BatchNo);
                            ws.Cells[row, 4].PutValue(projectBatch.Scope);
                            ws.Cells[row, 5].PutValue(Convert.ToDateTime(projectBatch.ReceivedDate).ToString("dd-MMM-yyyy"));
                            if (projectBatch.PlannedDeliveryDate != null)
                                ws.Cells[row, 6].PutValue(Convert.ToDateTime(projectBatch.PlannedDeliveryDate).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 7].PutValue(projectBatch.InputCount);
                            ws.Cells[row, 8].PutValue(projectBatch.InputCountType);
                            ws.Cells[row, 9].PutValue(projectBatch.Status);
                            if (Status == "D")
                            {
                                ws.Cells[row, 10].PutValue(Convert.ToDateTime(projectBatch.DeliveredDate).ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 11].PutValue(projectBatch.DeliveredCount);
                            }
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 4].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                            if (Status == "D")
                            {
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                            }
                            #endregion
                            row++;
                        }
                        #endregion

                        #region Saving and downloading the report
                        ws.AutoFitColumns();
                        string tempPath = System.IO.Path.GetTempPath();
                        string filename = tempPath + FileName;
                        wb.Save(filename);

                        //Read the file into a Byte Array.
                        byte[] bytes = File.ReadAllBytes(filename);

                        //Set the Response Content.
                        response.Content = new ByteArrayContent(bytes);

                        //Set the Response Content Length.
                        response.Content.Headers.ContentLength = bytes.LongLength;

                        //Set the Content Disposition Header Value and FileName.
                        response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                        response.Content.Headers.ContentDisposition.FileName = FileName;

                        //Set the File Content Type.
                        response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(filename));

                        return response;
                        #endregion
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No data found");
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Update Project Batch as Delivered
        [HttpPost]
        [Route("DeliverProjectBatch")]
        public HttpResponseMessage DeliverProjectBatch([FromBody]ProjectBatchStatus projectBatchStatus)
        {
            try
            {
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Batch Delivery Details data");

                if (!AccessControl.CanUserAccessPage(projectBatchStatus.UserID, "Deliver Project Batch"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateProjectBatchDeliveryDetails";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", projectBatchStatus.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectBatchStatus.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", projectBatchStatus.BatchNo);
                    cmd.Parameters.AddWithValue("@DeliveredDate", projectBatchStatus.DeliveredDate);
                    cmd.Parameters.AddWithValue("@DeliveredCount", projectBatchStatus.DeliveredCount);
                    cmd.Parameters.AddWithValue("@UserID", projectBatchStatus.UserID);
                    #endregion

                    //Calling sp to update project batch delivery details
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                    {
                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Project Batch Delivered Details
        [HttpGet]
        [Route("ReadProjectBatchDeliveredDetails/{CustomerCode}/{ProjectCode}/{BatchNo}/{UserID}")]
        public IHttpActionResult ReadProjectBatchDeliveredDetails(string CustomerCode, string ProjectCode, string BatchNo, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Deliver Project Batch"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                ProjectBatchStatus projectBatchStatus = new ProjectBatchStatus();       //Create a project batch status instance
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spReadProjectBatchDeliveryDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    #endregion

                    //Calling sp to get project batch status details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        projectBatchStatus.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectBatchStatus.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectBatchStatus.BatchNo = sqlReader["BatchNo"].ToString();
                        projectBatchStatus.DeliveredDate = Convert.ToDateTime(sqlReader["DeliveredDate"]);
                        projectBatchStatus.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        conn.Close();

                        //return project batch status to the request
                        return Ok(projectBatchStatus);
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Delete Delivered Project Batch
        [HttpPatch]
        [Route("DeleteDeliveredProjectBatch/{CustomerCode}/{ProjectCode}/{BatchNo}/{UserID}")]
        public HttpResponseMessage DeleteDeliveredProjectBatch(string CustomerCode, string ProjectCode, string BatchNo, string UserID)
        {
            try
            {
                //No additional page is added here, user who can deliver the batch, s/he can delete too
                if (!AccessControl.CanUserAccessPage(UserID, "Deliver Project Batch"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spDeleteDeliveredProjectBatchDetails";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to delete delivered Project Batch
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.ToLower() == "deleted")
                    {
                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
    }
}
