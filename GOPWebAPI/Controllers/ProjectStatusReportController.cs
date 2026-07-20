using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Xml;
using System.Xml.Serialization;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/ProjectStatusReport")]
    public class ProjectStatusReportController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public ProjectStatusReportController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Project Codes of Customer
        [HttpGet]
        [Route("ReadProjectCodesOfCustomer/{CustomerCode}")]
        public IHttpActionResult ReadProjectCodesOfCustomer(string CustomerCode)
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> ProjectCodeList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectCodesOfCustomer";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);

                    //Calling sp to get list of project codes
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CustomerCodeProjectCodeBatchNo projectCode = new CustomerCodeProjectCodeBatchNo();

                        projectCode.CustomerCode = CustomerCode;
                        projectCode.ProjectCode = Convert.ToString(sqlReader["ProjectCode"]);
                        projectCode.ProjectInputCount = Convert.ToInt64(sqlReader["InputCount"]);
                        ProjectCodeList.Add(projectCode);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectCodeList);
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

        #region Read Batches of Project
        [HttpGet]
        [Route("ReadBatchesOfProject/{CustomerCode}/{ProjectCode}")]
        public IHttpActionResult ReadBatchesOfProject(string CustomerCode, string ProjectCode)
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> BatchNoList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spBatchNosOfProject";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);

                    //Calling sp to get list of project batch nos.
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CustomerCodeProjectCodeBatchNo batchNo = new CustomerCodeProjectCodeBatchNo();

                        batchNo.CustomerCode = CustomerCode;
                        batchNo.ProjectCode = ProjectCode;
                        batchNo.BatchNo = Convert.ToString(sqlReader["BatchNo"]);

                        BatchNoList.Add(batchNo);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(BatchNoList);
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

        #region Read Project Details
        [HttpGet]
        [Route("ReadProjectDetails/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadProjectDetails(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                ProjectDetailsModel objProjectDetails = new ProjectDetailsModel();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get project details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        objProjectDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        objProjectDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        objProjectDetails.BatchNo = Convert.ToString(sqlReader["BatchNo"]);
                        objProjectDetails.Scope = Convert.ToString(sqlReader["Scope"]);
                        objProjectDetails.InputCount = Convert.ToInt64(sqlReader["InputCount"]);
                        objProjectDetails.ReceivedOn = Convert.ToDateTime(sqlReader["ReceivedOn"]);
                        objProjectDetails.DeliveredOn = sqlReader["DeliveredOn"] == DBNull.Value ? null : (DateTime?)sqlReader["DeliveredOn"];
                        conn.Close();

                        //return scope to the request
                        return Ok(objProjectDetails);
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

        #region Read Project Status Report data
        [HttpGet]
        [Route("ReadProjectStatusReportData/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadProjectStatusReportData(string CustomerCode, string ProjectCode,string BatchNo = "")
        {
            try
            {
                int SlNo = 1;

                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the Project Status Report Data
                List<ProjectStatusReportModel> ProjectStatusList = new List<ProjectStatusReportModel>();

                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectStatusReport";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of Project Status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectStatusReportModel projectStatusReportModel = new ProjectStatusReportModel();

                        projectStatusReportModel.SlNo = SlNo;
                        projectStatusReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                        projectStatusReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                        projectStatusReportModel.Activity = sqlReader["Activity"].ToString();
                        projectStatusReportModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        projectStatusReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        projectStatusReportModel.ProductionPendingCount = Convert.ToInt32(sqlReader["ProductionPendingCount"]);
                        projectStatusReportModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                        projectStatusReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        projectStatusReportModel.QCPendingCount = Convert.ToInt32(sqlReader["QCPendingCount"]);
                        projectStatusReportModel.LastUpdatedDate = sqlReader["LastUpdatedDate"] == DBNull.Value ? null : (DateTime?)sqlReader["LastUpdatedDate"];

                        ProjectStatusList.Add(projectStatusReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectStatusList);
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

        #region Export Project Status Report to Excel
        [HttpGet]
        [Route("ExportProjectStatusReportToExcel/{CustomerCode}/{ProjectCode}/{BatchNo?}/{UserID?}")]
        public HttpResponseMessage ExportProjectStatusReportToExcel(string CustomerCode, string ProjectCode, string BatchNo = "", string UserID = "")
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Project Status Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = string.Empty;

                    //Swagger fix
                    if (BatchNo == "{BatchNo}")
                        BatchNo = "";

                    if (string.IsNullOrEmpty(BatchNo))
                        FileName = "ProjectStatusReport_" + CustomerCode + '_' + ProjectCode + ".xlsx";
                    else
                        FileName = "ProjectStatusReport_" + CustomerCode + '_' + ProjectCode + '_' + BatchNo + ".xlsx";

                    //Create a list to hold the Project Status Report Data
                    List<ProjectStatusReportModel> ProjectStatusList = new List<ProjectStatusReportModel>();

                    System.Data.Common.DbDataReader sqlReader;
                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    var ws = wb.Worksheets[0];
                    int row = 6, counter = 1;

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignDataWithoutBorder = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleRightAlignData = ws.Cells[0, 0].GetStyle();

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

                    styleCenterAlignDataWithoutBorder.HorizontalAlignment = TextAlignmentType.Center;

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
                        cmd.CommandText = "spProjectStatusReport";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                        //Calling sp to get uploaded data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                ProjectStatusReportModel projectStatusReportModel = new ProjectStatusReportModel();

                                projectStatusReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                                projectStatusReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                                projectStatusReportModel.Activity = sqlReader["Activity"].ToString();
                                projectStatusReportModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                                projectStatusReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                projectStatusReportModel.ProductionPendingCount = Convert.ToInt32(sqlReader["ProductionPendingCount"]);
                                projectStatusReportModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                                projectStatusReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                                projectStatusReportModel.QCPendingCount = Convert.ToInt32(sqlReader["QCPendingCount"]);
                                projectStatusReportModel.LastUpdatedDate = sqlReader["LastUpdatedDate"] == DBNull.Value ? null : (DateTime?)sqlReader["LastUpdatedDate"];

                                ProjectStatusList.Add(projectStatusReportModel);
                            }
                            #endregion
                            conn.Close();

                            #region Writing report header
                            for (int col = 0; col <= 10; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 11);
                            ws.Cells[1, 0].PutValue("Projects Status Report");

                            for (int col = 0; col <= 1; col++)
                                ws.Cells[3, col].SetStyle(styleCenterAlignDataWithoutBorder);
                            ws.Cells.Merge(3, 0, 1, 2);
                            ws.Cells[3, 0].PutValue("Customer Code: " + CustomerCode);

                            for (int col = 5; col <= 6; col++)
                                ws.Cells[3, col].SetStyle(styleCenterAlignDataWithoutBorder);
                            ws.Cells.Merge(3, 5, 1, 2);
                            ws.Cells[3, 5].PutValue("Project Code: " + ProjectCode);

                            if (!string.IsNullOrEmpty(BatchNo))
                            {
                                for (int col = 9; col <= 10; col++)
                                    ws.Cells[3, col].SetStyle(styleCenterAlignDataWithoutBorder);
                                ws.Cells.Merge(3, 9, 1, 2);
                                ws.Cells[3, 9].PutValue("Batch No.: " + BatchNo);
                            }
                            #endregion

                            #region Writing column headings
                            ws.Cells[5, 0].PutValue("S.No.");
                            ws.Cells[5, 1].PutValue("Employee Code");
                            ws.Cells[5, 2].PutValue("Employee Name");
                            ws.Cells[5, 3].PutValue("Activity");
                            ws.Cells[5, 4].PutValue("Production Allocated Count");
                            ws.Cells[5, 5].PutValue("Production Completed Count");
                            ws.Cells[5, 6].PutValue("Production Pending Count");
                            ws.Cells[5, 7].PutValue("QC Allocated Count");
                            ws.Cells[5, 8].PutValue("QC Completed Count");
                            ws.Cells[5, 9].PutValue("QC Pending Count");
                            ws.Cells[5, 10].PutValue("Last Updated Date");

                            for (int c = 0; c <= 10; c++)
                                ws.Cells[5, c].SetStyle(styleHeader);
                            #endregion

                            foreach (ProjectStatusReportModel projectStatusReportModel in ProjectStatusList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(counter);
                                ws.Cells[row, 1].PutValue(projectStatusReportModel.EmployeeCode);
                                ws.Cells[row, 2].PutValue(projectStatusReportModel.EmployeeName);
                                ws.Cells[row, 3].PutValue(projectStatusReportModel.Activity);
                                ws.Cells[row, 4].PutValue(projectStatusReportModel.ProductionAllocatedCount);
                                ws.Cells[row, 5].PutValue(projectStatusReportModel.ProductionCompletedCount);
                                ws.Cells[row, 6].PutValue(projectStatusReportModel.ProductionPendingCount);
                                ws.Cells[row, 7].PutValue(projectStatusReportModel.QCAllocatedCount);
                                ws.Cells[row, 8].PutValue(projectStatusReportModel.QCCompletedCount);
                                ws.Cells[row, 9].PutValue(projectStatusReportModel.QCPendingCount);
                                if (projectStatusReportModel.LastUpdatedDate != null)
                                    ws.Cells[row, 10].PutValue(Convert.ToDateTime(projectStatusReportModel.LastUpdatedDate).ToString("dd-MMM-yyyy"));
                                #endregion

                                #region setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                #endregion

                                row++;
                                counter++;
                            }

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
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Project Status Activity Summary
        [HttpGet]
        [Route("ReadProjectStatusActivitySummary/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadProjectStatusActivitySummary(string CustomerCode, string ProjectCode,string BatchNo = "")
        {
            try
            {
                int index = 1;

                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the Project Status Activity Summary Data
                List<ProjectStatusActivitySummaryModel> ProjectStatusActivityList = new List<ProjectStatusActivitySummaryModel>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectStatusActivitywiseSummary";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of Project Status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectStatusActivitySummaryModel projectStatusActivitySummaryModel = new ProjectStatusActivitySummaryModel();

                        projectStatusActivitySummaryModel.index = index;
                        projectStatusActivitySummaryModel.Activity = sqlReader["Activity"].ToString();
                        projectStatusActivitySummaryModel.ActivityCount = Convert.ToInt32(sqlReader["ActivityCount"]);
                        projectStatusActivitySummaryModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        projectStatusActivitySummaryModel.ProductionAllocatedPercentage = Convert.ToDecimal(sqlReader["ProductionAllocatedPercentage"]);
                        projectStatusActivitySummaryModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        projectStatusActivitySummaryModel.ProductionCompletedPercentage = Convert.ToDecimal(sqlReader["ProductionCompletedPercentage"]);
                        projectStatusActivitySummaryModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                        projectStatusActivitySummaryModel.QCAllocatedPercentage = Convert.ToDecimal(sqlReader["QCAllocatedPercentage"]);
                        projectStatusActivitySummaryModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        projectStatusActivitySummaryModel.QCCompletedPercentage = Convert.ToDecimal(sqlReader["QCCompletedPercentage"]);

                        ProjectStatusActivityList.Add(projectStatusActivitySummaryModel);
                        index++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectStatusActivityList);
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
    }
}
