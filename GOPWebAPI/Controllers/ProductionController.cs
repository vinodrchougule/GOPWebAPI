using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;
using System.Xml;
using System.Xml.Serialization;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using Microsoft.Ajax.Utilities;
using Newtonsoft.Json.Linq;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/production")]
    public class ProductionController : ApiController
    {
        private BLLProduction _BLLProduction;
        private BLLAccessControl _BLLAccessControl;
        public ProductionController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLProduction = new BLLProduction(connectionString);
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Projects Customer Codes of User
        [HttpGet]
        [Route("ReadProjectCustomerCodesOfUser/{ProductionUser}/{status?}")]
        public IHttpActionResult ReadProjectCustomerCodesOfUser(string ProductionUser, string Status = "O")
        {
            try
            {
                //if (!AccessControl.CanUserAccessPage(ProductionUser, "Production Download/Upload"))
                //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                List<Customer> CustomerCodeList = new List<Customer>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionGetCustomerCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Status", Status);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

                    //Calling sp to get list of Customer Codes
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Customer c = new Customer();
                        c.CustomerCode = sqlReader["CustomerCode"].ToString();
                        CustomerCodeList.Add(c);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(CustomerCodeList);
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

        #region Read Customer Project Codes of User
        [HttpGet]
        [Route("ReadCustomerProjectCodesOfUser/{CustomerCode}/{ProductionUser}/{status?}")]
        public IHttpActionResult ReadCustomerProjectCodesOfUser(string CustomerCode, string ProductionUser, string Status = "O")
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> ProjectCodeList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionGetProjectCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);
                    cmd.Parameters.AddWithValue("@Status", Status);


                    //Calling sp to get list of projects
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CustomerCodeProjectCodeBatchNo projectCode = new CustomerCodeProjectCodeBatchNo();

                        projectCode.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectCode.ProjectCode = Convert.ToString(sqlReader["ProjectCode"]);
                        ProjectCodeList.Add(projectCode);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectCodeList.Distinct());
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

        #region Read Customer Project Batch Nos. of User
        [HttpGet]
        [Route("ReadCustomerProjectBatchNosOfUser/{CustomerCode}/{ProjectCode}/{ProductionUser}/{status?}")]
        public IHttpActionResult ReadCustomerProjectBatchNosOfUser(string CustomerCode, string ProjectCode, string ProductionUser, string Status = "O")
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> BatchNoList = new List<CustomerCodeProjectCodeBatchNo>();
                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionGetBatchNos";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);
                    cmd.Parameters.AddWithValue("@Status", Status);

                    //Calling sp to get list of projects
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CustomerCodeProjectCodeBatchNo batchNo = new CustomerCodeProjectCodeBatchNo();

                        batchNo.CustomerCode = sqlReader["CustomerCode"].ToString();
                        batchNo.ProjectCode = Convert.ToString(sqlReader["ProjectCode"]);
                        batchNo.BatchNo = Convert.ToString(sqlReader["BatchNo"]);

                        BatchNoList.Add(batchNo);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(BatchNoList.Distinct());
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

        #region Read Project Activities of User
        [HttpGet]
        [Route("ReadProjectActivitiesOfUser/{CustomerCode}/{ProjectCode}/{ProductionUser}/{BatchNo?}")]
        public IHttpActionResult ReadProjectActivitiesOfUser(string CustomerCode, string ProjectCode, string ProductionUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Activity Counts
                List<ProjectActivitiesCount> ProjectActivitiesCountList = new List<ProjectActivitiesCount>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionProjectActivityCountStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

                    //Calling sp to get list of Activity Counts
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectActivitiesCount pac = new ProjectActivitiesCount();

                        pac.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        pac.Activity = sqlReader["Activities"].ToString();
                        pac.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        pac.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        pac.ProductionPendingCount = Convert.ToInt32(sqlReader["ProductionPendingCount"]);
                        pac.ProductionErrorCount = Convert.ToInt32(sqlReader["ProductionErrorCount"]);
                        pac.IsAllocationDownloadedForProduction = Convert.ToInt32(sqlReader["IsAllocationDownloadedForProduction"]);
                        pac.IsProductionErrorDownloaded = Convert.ToInt32(sqlReader["IsProductionErrorDownloaded"]);

                        ProjectActivitiesCountList.Add(pac);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectActivitiesCountList);
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

        #region Production Download
        #region Download Production Allocation of User
        [HttpPost]
        [Route("DownloadProductionAllocationOfUser")]
        public HttpResponseMessage DownloadProductionAllocationOfUser([FromBody] ProductionDownload productionDownload)
        {
            try
            {
                string FileName = "ProductionAllocation_" + productionDownload.ProductionUser + '_' + productionDownload.ProductionAllocationID.ToString() + ".xlsx";

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUpdateDownload";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionDownload.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@UserID", productionDownload.ProductionUser.Trim());

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "updated")
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/"));
                        string AllocatedFileName = Path.Combine(dirUploads.FullName, arrResult[1].ToString());
                        int ProductionUserColumnNo = -1, dRowNo = 0;

                        if (!File.Exists(AllocatedFileName))
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocation file not found");

                        #region Copy the Project Allocated User data from allocated file to new file and download the file
                        Workbook wbAF = new Workbook();         //Allocated file
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        //wbAF.Open(AllocatedFileName);
                        wbAF.LoadData(AllocatedFileName);
                        var ws = wbAF.Worksheets[0];            //source worksheet
                        //Cells cells = ws.Cells;

                        Workbook wbUDF = new Workbook();         //User Download file
                        Aspose.Cells.License l1 = new Aspose.Cells.License();
                        l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        //wbUDF.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                        wbUDF.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                        var wd = wbUDF.Worksheets[0];           //destination worksheet

                        #region Copying header as is from allocated file
                        for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                        {
                            wd.Cells[0, sCol].Copy(ws.Cells[0, sCol]);
                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "production user")
                                ProductionUserColumnNo = sCol;
                        }
                        #endregion

                        #region Copying data rows of user from allocated file
                        dRowNo = 1;
                        for (int sRow = 1; sRow <= ws.Cells.MaxRow; sRow++)
                        {
                            if (ws.Cells[sRow, ProductionUserColumnNo].StringValue.Trim().ToLower() == productionDownload.ProductionUser.Trim().ToLower())
                            {
                                //wd.Cells.CopyRow(cells, sRow, dRowNo);
                                for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                                    wd.Cells[dRowNo, sCol].Copy(ws.Cells[sRow, sCol]);
                                dRowNo++;
                            }
                        }
                        #endregion

                        #region Write Status as 'Production Downloaded' for all rows and adding 'Production Comments' column just heading 
                        int StatusColNo = wd.Cells.MaxColumn + 1;
                        int ProductionCommentsColNo = StatusColNo + 1;
                        bool IsProductionCommentsColumnExists = false;
                        wd.Cells[0, StatusColNo].PutValue("Status");
                        for (int col = 0; col <= wd.Cells.MaxColumn; col++)
                        {
                            if (wd.Cells[0, col].StringValue.Trim().ToLower() == "production comments")
                            {
                                IsProductionCommentsColumnExists = true;
                                break;
                            }
                        }
                        if (!IsProductionCommentsColumnExists)
                            wd.Cells[0, ProductionCommentsColNo].PutValue("Production Comments");
                        for (int dRow = 1; dRow <= wd.Cells.MaxRow; dRow++)
                            wd.Cells[dRow, StatusColNo].PutValue("Production Downloaded");
                        #endregion

                        #region Save and Download file
                        wd.AutoFitColumns();
                        string tempPath = System.IO.Path.GetTempPath();
                        string filename = tempPath + FileName;
                        wbUDF.Save(filename);

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
                        #endregion
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
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
        #endregion

        #region Production Error Download
        [HttpPost]
        [Route("DownloadProductionAllocationErrorsOfUser")]
        public HttpResponseMessage DownloadProductionAllocationErrorsOfUser([FromBody] ProductionDownload productionDownload)
        {
            try
            {
                string FileName = "ProductionErrors_" + productionDownload.ProductionUser + '_' + productionDownload.ProductionAllocationID.ToString() + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionErrorDownload";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionDownload.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionDownload.ProductionUser.Trim());

                    //Calling sp to get Allocated data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                            ws.Cells[0, col].PutValue(sqlReader.GetName(col));

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                                ws.Cells[row, col].PutValue(sqlReader.GetValue(col));
                            ++row;
                        }

                        int IsDeletedColNo = 0;
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            if (ws.Cells[0, col].StringValue.ToLower() == "isdeleted")
                            {
                                IsDeletedColNo = col;
                                break;
                            }
                        }

                        ws.Cells.DeleteRange(0, IsDeletedColNo, ws.Cells.MaxRow, IsDeletedColNo, ShiftType.Left);

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

        #region Production Error Upload
        [HttpPost]
        [Route("ValidateAndUploadProductionErrorSKUs")]
        public HttpResponseMessage ValidateAndUploadProductionErrorSKUs([FromBody] ProductionUpload productionUpload)
        {
            try
            {
                string AllocatedFileName = string.Empty;
                string AllocatedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                string AllocatedFileFullName = string.Empty;
                System.Data.Common.DbDataReader sqlReader;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Production Error Upload Data");

                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + productionUpload.UploadedFileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload Production Error Upload file");
                #endregion

                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + productionUpload.UploadedFileName);
                string UploadedFileExtension = Path.GetExtension(UploadedFilepath);
                int ActivitiesColumnNo = -1, ProductionUserColumnNo = -1, StatusColumnNo = -1;

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);


                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.LoadData(UploadedFilepath);
                Worksheet wsUF = wbUF.Worksheets[0];

                #region Check uploaded file has data rows
                if (wsUF.Cells.MaxRow <= 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Error file has no data rows in first worksheet.");
                #endregion

                #region Check Activities, Production User,Production Comments, Status columns exists, if exists get column Nos. 'Production Comments' column is mandatory to avoid inconsistency
                int ProductionCommentsColumnNo = -1;
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    if (string.IsNullOrEmpty(wsUF.Cells[0, col].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Blank/Empty column heading in uploaded file. Column No.: " + (col + 1).ToString());

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "activities")
                        ActivitiesColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "production user")
                        ProductionUserColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "status")
                        StatusColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "production comments")
                        ProductionCommentsColumnNo = col;

                    if (ActivitiesColumnNo >= 0 && ProductionUserColumnNo >= 0 && StatusColumnNo >= 0 && ProductionCommentsColumnNo >= 0)
                        break;
                }

                if (ActivitiesColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column not found in uploaded production file");

                if (ProductionUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column not found in uploaded production file");

                if (StatusColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Status' column not found in uploaded production file");

                if (ProductionCommentsColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production Comments' column not found in uploaded production file");
                #endregion

                #region Check production upload file should have production allocated user entries only
                for (int row = 1; row <= wsUF.Cells.MaxRow; row++)
                {
                    //if (wsUF.Cells[row, ActivitiesColumnNo].StringValue.Trim().ToLower() != productionUpload.Activities.Trim().ToLower())
                    //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Error file should have only selected allocation activities");

                    if (wsUF.Cells[row, ProductionUserColumnNo].StringValue.Trim().ToLower() != productionUpload.UserID.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Error file should have only current production user rows");

                    if (wsUF.Cells[row, StatusColumnNo].StringValue.Trim().ToLower() != "production completed")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Error file all SKUs should have status as 'Production Completed'");
                }
                #endregion

                #region Get the allocated file name
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

                //Initialize command object
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "spProductionAllocationGetUploadedFileNameByID";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                //Add parameters with values
                cmd.Parameters.AddWithValue("@ProductionAllocationID", productionUpload.ProductionAllocationID);

                //Calling sp to get Allocated FileName
                conn.Open();
                cmd.CommandTimeout = 0;
                sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                if (sqlReader.Read())
                {
                    AllocatedFileName = sqlReader["AllocatedFileName"].ToString();
                    AllocatedFileFullName = Path.Combine(AllocatedFilePath, AllocatedFileName);
                }
                conn.Close();

                //Check whether File exists.
                if (!File.Exists(AllocatedFileFullName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocated File not found");
                #endregion

                #region Check all columns from uploaded file exists in allocated file
                Workbook wbAF = new Workbook();                       //Allocated file work book
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbAF.LoadData(AllocatedFileFullName);
                Worksheet wsAF = wbAF.Worksheets[0];

                string UploadedFileColumnName = string.Empty;
                bool IsUploadedColumnExists = false;
                for (int colUF = 0; colUF <= wsUF.Cells.MaxColumn; colUF++)
                {
                    UploadedFileColumnName = wsUF.Cells[0, colUF].StringValue.Trim();
                    if (UploadedFileColumnName.Trim().ToLower() != "production comments" && UploadedFileColumnName.Trim().ToLower() != "status" && UploadedFileColumnName.Trim().ToLower() != "qc user" && UploadedFileColumnName.Trim().ToLower() != "qc comments")
                    {
                        IsUploadedColumnExists = false;
                        for (int colAF = 0; colAF <= wsAF.Cells.MaxColumn; colAF++)
                        {
                            if (wsAF.Cells[0, colAF].StringValue.Trim().ToLower() == UploadedFileColumnName.Trim().ToLower())
                            {
                                IsUploadedColumnExists = true;
                                break;
                            }
                        }

                        if (!IsUploadedColumnExists)
                            break;
                    }
                }

                if (!IsUploadedColumnExists)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, UploadedFileColumnName + " - uploaded column not found. Columns in uploaded production file should be as same as allocated file");
                #endregion

                #region Check all columns from allocated file exists in uploaded file
                string AllocatedFileColumnName = string.Empty;
                bool IsAllocatedColumnExists = false;
                for (int colAF = 0; colAF <= wsAF.Cells.MaxColumn; colAF++)
                {
                    AllocatedFileColumnName = wsAF.Cells[0, colAF].StringValue.Trim();
                    IsAllocatedColumnExists = false;
                    for (int colUF = 0; colUF <= wsUF.Cells.MaxColumn; colUF++)
                    {
                        if (wsUF.Cells[0, colUF].StringValue.Trim().ToLower() == AllocatedFileColumnName.Trim().ToLower())
                        {
                            IsAllocatedColumnExists = true;
                            break;
                        }
                    }

                    if (!IsAllocatedColumnExists)
                        break;
                }

                if (!IsAllocatedColumnExists)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, AllocatedFileColumnName + " - allocated column not found. Columns in uploaded production file should be as same as allocated file");
                #endregion

                #region Form string to create temp table
                string sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                #region Generate dynamic column name and data type string
                string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    if (!string.IsNullOrEmpty(wsUF.Cells[0, col].StringValue.Trim()))
                    {
                        ColumnName = "[" + wsUF.Cells[0, col].StringValue.Trim() + "]";
                        ColumnDataType = "VARCHAR(4000)";

                        if (string.IsNullOrEmpty(ColumnNameDataTypeString.Trim()))
                            ColumnNameDataTypeString = ColumnName + " " + ColumnDataType;
                        else
                            ColumnNameDataTypeString += "," + ColumnName + " " + ColumnDataType;
                    }
                }
                #endregion

                ColumnNameDataTypeString += ",IsDeleted BIT";

                strSQL += ColumnNameDataTypeString + ")";

                #endregion

                #region Create temp table, write uploaded file data, upload Production, Move file
                #region Creating SQL table
                SqlCommand cmdCreateSQLTable = conn.CreateCommand();
                cmdCreateSQLTable.CommandText = strSQL;
                cmdCreateSQLTable.CommandType = System.Data.CommandType.Text;
                conn.Open();
                cmd.CommandTimeout = 0;
                cmdCreateSQLTable.ExecuteNonQuery();
                #endregion

                #region Writing data to SQL table
                DataFormatConverter dataFormatConverter = new DataFormatConverter();
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion

                #region Create Production Error file Upload
                //Initialize command object
                SqlCommand cmdPA = new SqlCommand();
                cmdPA.Connection = conn;
                cmdPA.CommandType = CommandType.StoredProcedure;
                cmdPA.CommandText = "spProductionErrorValidateAndUpload";

                #region Adding Stored Procedure Parameters
                cmdPA.Parameters.AddWithValue("@ProductionAllocationID", productionUpload.ProductionAllocationID);
                cmdPA.Parameters.AddWithValue("@UploadedFileExtension", UploadedFileExtension);
                cmdPA.Parameters.AddWithValue("@TempTableName", sqlTableName);
                cmdPA.Parameters.AddWithValue("@UserID", productionUpload.UserID);
                #endregion

                //Calling sp to create production upload
                cmdPA.CommandTimeout = 0;
                string Result = cmdPA.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split(',');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    long ProductionUploadID = Convert.ToInt64(arrResult[1]);
                    string NewUploadedFileName = arrResult[2];

                    #region Input File move Starts
                    if (File.Exists(dirTemp + productionUpload.UploadedFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionUpload/"));
                        FileOperations.MoveFile(dirTemp, productionUpload.UploadedFileName, dirUploads, NewUploadedFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, ProductionUploadID);
                }
                else
                {
                    //return error response status code
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
                }
                #endregion
                #endregion
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Production Upload
        [HttpPost]
        [Route("ValidateAndUploadProduction")]
        public HttpResponseMessage ValidateAndUploadProduction([FromBody] ProductionUpload productionUpload)
        {
            try
            {
                string AllocatedFileName = string.Empty;
                string AllocatedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                string AllocatedFileFullName = string.Empty;
                System.Data.Common.DbDataReader sqlReader;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Production Upload Data");

                if (!AccessControl.CanUserAccessPage(productionUpload.UserID, "Production Download-Upload"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + productionUpload.UploadedFileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload Production Upload file");
                #endregion

                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + productionUpload.UploadedFileName);
                string UploadedFileExtension = Path.GetExtension(UploadedFilepath);
                int ActivitiesColumnNo = -1, ProductionUserColumnNo = -1, StatusColumnNo = -1;

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.LoadData(UploadedFilepath);
                Worksheet wsUF = wbUF.Worksheets[0];

                #region Check uploaded file has data rows
                if (wsUF.Cells.MaxRow <= 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production file has no data rows in first worksheet.");
                #endregion

                #region Check uploaded file for duplicate columns if any
                for (int mcol = 0; mcol < wsUF.Cells.MaxColumn; mcol++)              //main column
                {
                    for (int ncol = mcol + 1; ncol <= wsUF.Cells.MaxColumn; ncol++)      //navigation column
                    {
                        if (wsUF.Cells[0, mcol].StringValue.Trim().ToLower() == wsUF.Cells[0, ncol].StringValue.Trim().ToLower())
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file has duplicate columns. column name:" + wsUF.Cells[0, mcol].StringValue.Trim());
                    }
                }
                #endregion

                #region Check any value exceeds 4000 characters from all columns
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    for (int row = 1; row <= wsUF.Cells.MaxRow; row++)
                    {
                        if (wsUF.Cells[row, col].StringValue.Length > 4000)
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'" + wsUF.Cells[0, col].StringValue + "' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Check Activities, Production User,Production Comments, Status columns exists, if exists get column Nos. 'Production Comments' column is mandatory to avoid inconsistency
                int ProductionCommentsColumnNo = -1;
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    if (string.IsNullOrEmpty(wsUF.Cells[0, col].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Blank/Empty column heading in uploaded file. Column No.: " + (col + 1).ToString());

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "activities")
                        ActivitiesColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "production user")
                        ProductionUserColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "status")
                        StatusColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "production comments")
                        ProductionCommentsColumnNo = col;

                    if (ActivitiesColumnNo >= 0 && ProductionUserColumnNo >= 0 && StatusColumnNo >= 0 && ProductionCommentsColumnNo >= 0)
                        break;
                }

                if (ActivitiesColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column not found in uploaded production file");

                if (ProductionUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column not found in uploaded production file");

                if (StatusColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Status' column not found in uploaded production file");

                if (ProductionCommentsColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production Comments' column not found in uploaded production file");
                #endregion

                #region Check production upload file should have only selected activities and production allocated user entries only
                for (int row = 1; row <= wsUF.Cells.MaxRow; row++)
                {
                    if (wsUF.Cells[row, ProductionUserColumnNo].StringValue.Trim().ToLower() != productionUpload.UserID.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file should have only current production user rows");

                    if (wsUF.Cells[row, StatusColumnNo].StringValue.Trim().ToLower() != "production completed")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file all SKUs should have status as 'Production Completed'");
                }
                #endregion

                #region Get the allocated file name
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

                //Initialize command object
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "spProductionAllocationGetUploadedFileNameByID";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                //Add parameters with values
                cmd.Parameters.AddWithValue("@ProductionAllocationID", productionUpload.ProductionAllocationID);

                //Calling sp to get Allocated FileName
                conn.Open();
                cmd.CommandTimeout = 0;
                sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                if (sqlReader.Read())
                {
                    AllocatedFileName = sqlReader["AllocatedFileName"].ToString();
                    AllocatedFileFullName = Path.Combine(AllocatedFilePath, AllocatedFileName);
                }
                conn.Close();

                //Check whether File exists.
                if (!File.Exists(AllocatedFileFullName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocated File not found");
                #endregion

                #region Check all columns from uploaded file exists in allocated file
                Workbook wbAF = new Workbook();                       //Allocated file work book
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbAF.LoadData(AllocatedFileFullName);
                Worksheet wsAF = wbAF.Worksheets[0];

                string UploadedFileColumnName = string.Empty;
                bool IsUploadedColumnExists = false;
                for (int colUF = 0; colUF <= wsUF.Cells.MaxColumn; colUF++)
                {
                    UploadedFileColumnName = wsUF.Cells[0, colUF].StringValue.Trim();
                    if (UploadedFileColumnName.Trim().ToLower() != "production comments" && UploadedFileColumnName.Trim().ToLower() != "status")
                    {
                        IsUploadedColumnExists = false;
                        for (int colAF = 0; colAF <= wsAF.Cells.MaxColumn; colAF++)
                        {
                            if (wsAF.Cells[0, colAF].StringValue.Trim().ToLower() == UploadedFileColumnName.Trim().ToLower())
                            {
                                IsUploadedColumnExists = true;
                                break;
                            }
                        }

                        if (!IsUploadedColumnExists)
                            break;
                    }
                }

                if (!IsUploadedColumnExists)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, UploadedFileColumnName + " - uploaded column not found. Columns in uploaded production file should be as same as allocated file");
                #endregion

                #region Check all columns from allocated file exists in uploaded file
                string AllocatedFileColumnName = string.Empty;
                bool IsAllocatedColumnExists = false;
                for (int colAF = 0; colAF <= wsAF.Cells.MaxColumn; colAF++)
                {
                    AllocatedFileColumnName = wsAF.Cells[0, colAF].StringValue.Trim();
                    IsAllocatedColumnExists = false;
                    for (int colUF = 0; colUF <= wsUF.Cells.MaxColumn; colUF++)
                    {
                        if (wsUF.Cells[0, colUF].StringValue.Trim().ToLower() == AllocatedFileColumnName.Trim().ToLower())
                        {
                            IsAllocatedColumnExists = true;
                            break;
                        }
                    }

                    if (!IsAllocatedColumnExists)
                        break;
                }

                if (!IsAllocatedColumnExists)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, AllocatedFileColumnName + " - allocated column not found. Columns in uploaded production file should be as same as allocated file");
                #endregion

                #region Form string to create temp table
                string sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                #region Generate dynamic column name and data type string
                string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    if (!string.IsNullOrEmpty(wsUF.Cells[0, col].StringValue.Trim()))
                    {
                        ColumnName = "[" + wsUF.Cells[0, col].StringValue.Trim() + "]";
                        ColumnDataType = "VARCHAR(4000)";

                        if (string.IsNullOrEmpty(ColumnNameDataTypeString.Trim()))
                            ColumnNameDataTypeString = ColumnName + " " + ColumnDataType;
                        else
                            ColumnNameDataTypeString += "," + ColumnName + " " + ColumnDataType;
                    }
                }
                #endregion

                ColumnNameDataTypeString += ",IsDeleted BIT";

                strSQL += ColumnNameDataTypeString + ")";

                #endregion

                #region Create temp table, write uploaded file data, upload Production, Move file
                #region Creating SQL table
                SqlCommand cmdCreateSQLTable = conn.CreateCommand();
                cmdCreateSQLTable.CommandText = strSQL;
                cmdCreateSQLTable.CommandType = System.Data.CommandType.Text;
                conn.Open();
                cmdCreateSQLTable.CommandTimeout = 0;
                cmdCreateSQLTable.ExecuteNonQuery();
                #endregion

                #region Writing data to SQL table
                DataFormatConverter dataFormatConverter = new DataFormatConverter();
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion

                #region Create Production Upload
                //Initialize command object
                SqlCommand cmdPA = new SqlCommand();
                cmdPA.Connection = conn;
                cmdPA.CommandType = CommandType.StoredProcedure;
                cmdPA.CommandText = "spProductionValidateAndUpload";

                #region Adding Stored Procedure Parameters
                cmdPA.Parameters.AddWithValue("@ProductionAllocationID", productionUpload.ProductionAllocationID);
                cmdPA.Parameters.AddWithValue("@UploadedFileExtension", UploadedFileExtension);
                cmdPA.Parameters.AddWithValue("@TempTableName", sqlTableName);
                cmdPA.Parameters.AddWithValue("@WorkedHours", productionUpload.WorkedHours);
                cmdPA.Parameters.AddWithValue("@WorkedMinutes", productionUpload.WorkedMinutes);
                cmdPA.Parameters.AddWithValue("@UserID", productionUpload.UserID);
                #endregion

                //Calling sp to create production upload
                cmdPA.CommandTimeout = 0;
                string Result = cmdPA.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split(',');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    long ProductionUploadID = Convert.ToInt64(arrResult[1]);
                    string NewUploadedFileName = arrResult[2];

                    #region Input File move Starts
                    if (File.Exists(dirTemp + productionUpload.UploadedFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionUpload/"));
                        FileOperations.MoveFile(dirTemp, productionUpload.UploadedFileName, dirUploads, NewUploadedFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, ProductionUploadID);
                }
                else
                {
                    //return error response status code
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
                }

                #endregion
                #endregion
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region View Existing Production Uploads

        #region View Existing Production Uploads of Project By User
        [HttpGet]
        [Route("ReadExistingProjectUploadsByUser/{CustomerCode}/{ProjectCode}/{ProductionUser}/{BatchNo?}")]
        public IHttpActionResult ReadExistingProjectUploadsByUser(string CustomerCode, string ProjectCode, string ProductionUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Existing Production Uploads
                List<ProductionExistingUpload> productionExistingUploadList = new List<ProductionExistingUpload>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionGetExistingUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@UserID", ProductionUser);

                    //Calling sp to get list of existing Production Uploads
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionExistingUpload productionExistingUpload = new ProductionExistingUpload();

                        productionExistingUpload.ProductionUploadID = Convert.ToInt64(sqlReader["ProductionUploadID"]);
                        productionExistingUpload.UploadedOn = Convert.ToDateTime(sqlReader["UploadedOn"]);
                        productionExistingUpload.UploadedByUserName = sqlReader["UploadedByUserName"].ToString();
                        productionExistingUpload.Activities = sqlReader["Activities"].ToString();
                        productionExistingUpload.NoOfSKUs = Convert.ToInt32(sqlReader["NoOfSKUs"]);
                        productionExistingUpload.UploadedFileName = sqlReader["UploadedFileName"].ToString();
                        productionExistingUpload.IsProductionCompletedCountDownloaded = Convert.ToInt32(sqlReader["IsProductionCompletedCountDownloaded"]);

                        productionExistingUploadList.Add(productionExistingUpload);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(productionExistingUploadList);
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

        #region Delete Production Upload
        [HttpPatch]
        [Route("DeleteProductionUpload/{id}/{UserID}")]
        public HttpResponseMessage DeleteProductionUpload(long id, string UserID)
        {
            try
            {
                string UploadedFileName = string.Empty, fileFullName = string.Empty;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionDeleteUpload";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionUploadID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete Production Upload
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        UploadedFileName = arrResult[1];

                        #region Delete Uploaded File
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionUpload/");
                        fileFullName = Path.Combine(fileUploadedPath, UploadedFileName);
                        if (File.Exists(fileFullName))
                            File.Delete(fileFullName);
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

        #region Download Uploaded file by Production Upload ID
        [HttpGet]
        [Route("downloadfile/{id}")]
        public HttpResponseMessage DownloadUploadedFile(long id)
        {
            try
            {
                string UploadedFileName = string.Empty, fileFullName = string.Empty;
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionUpload/");

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionGetUploadedFileName";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionUploadID", id);

                    //Calling sp to get uploaded FileName
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        UploadedFileName = sqlReader["UploadedFileName"].ToString();
                        fileFullName = Path.Combine(fileUploadedPath, UploadedFileName);

                        //Check whether File exists.
                        if (!File.Exists(fileFullName))
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Production Uploaded File not found");

                        //Read the file into a Byte Array.
                        byte[] bytes = File.ReadAllBytes(fileFullName);

                        //Set the Response Content.
                        response.Content = new ByteArrayContent(bytes);

                        //Set the Response Content Length.
                        response.Content.Headers.ContentLength = bytes.LongLength;

                        //Set the Content Disposition Header Value and FileName.
                        response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                        response.Content.Headers.ContentDisposition.FileName = UploadedFileName;

                        //Set the File Content Type.
                        response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(fileFullName));

                        conn.Close();
                        return response;
                    }
                    else
                    {
                        conn.Close();
                        response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, "File not found");
                        return response;
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
        #endregion

        #region Read Production User Project Production Data for static data (delete after implementing dynamic)
        [HttpGet]
        [Route("ReadProductionUsersProjectProductionData/{ProductionAllocationID}/{UserID}")]
        public IHttpActionResult ReadProductionUsersProjectProductionData(long ProductionAllocationID, string UserID)
        {
            try
            {
                DataTable dt = null;
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersProjectProductionData";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", UserID);

                    //Calling sp to get updated data
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        dt = new DataTable();

                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            var column = new DataColumn(sqlReader.GetName(col));
                            dt.Columns.Add(column);
                        }

                        while (sqlReader.Read())
                        {
                            var r = dt.NewRow();
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                r[col] = sqlReader.GetValue(col);
                            }
                            dt.Rows.Add(r);
                        }
                    }
                    conn.Close();

                    return Ok(dt);
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

        #region Read Production Users Allocation Production Data
        [HttpGet]
        [Route("ReadProductionUsersAllocationProductionData/{ProductionAllocationID}/{UserID}")]
        public IHttpActionResult ReadProductionUsersAllocationProductionData(long ProductionAllocationID, string UserID)
        {
            try
            {
                DataTable dt = null;
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersAllocationProductionData";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", UserID);

                    //Calling sp to get updated data
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        dt = new DataTable();

                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            var column = new DataColumn(sqlReader.GetName(col));
                            dt.Columns.Add(column);
                        }

                        while (sqlReader.Read())
                        {
                            var r = dt.NewRow();
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                r[col] = sqlReader.GetValue(col);
                            }
                            dt.Rows.Add(r);
                        }
                    }
                    conn.Close();

                    return Ok(dt);
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

        #region Update Production Row Data
        [HttpPost]
        [Route("UpdateProductionRowData")]
        public HttpResponseMessage UpdateProductionRowData([FromBody] ProductionRowData ProductionRowData)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Production Row Data");

                #region Production Column and Row Values
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ProductionColumnNameAndValue>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, ProductionRowData.ProductionColumnValues);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateProductionRowData";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionRowData.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@UniqueColumnValue", ProductionRowData.UniqueColumnValue);
                    cmd.Parameters.Add(new SqlParameter("@ProductionColumnValues", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", ProductionRowData.UserID);
                    #endregion

                    //Calling sp to update project column read only status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
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

        #region Production Item Update
        [HttpPost]
        [Route("ProductionItemUpdate")]
        public HttpResponseMessage ProductionItemUpdate([FromBody] ProductionItem productionItem)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Production Row Data");

                #region Production Item Attribute
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ItemAttribute>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, productionItem.ItemAttributes);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionItemUpdate";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@ProductionItemID", productionItem.ProductionItemID);
                    cmd.Parameters.AddWithValue("@UniqueID", productionItem.UniqueID);
                    cmd.Parameters.AddWithValue("@ShortDescription", productionItem.ShortDescription);
                    cmd.Parameters.AddWithValue("@LongDescription", productionItem.LongDescription);
                    cmd.Parameters.AddWithValue("@UOM", productionItem.UOM);
                    cmd.Parameters.AddWithValue("@MFRName", productionItem.MFRName);
                    cmd.Parameters.AddWithValue("@MFRPN", productionItem.MFRPN);
                    cmd.Parameters.AddWithValue("@VendorName", productionItem.VendorName);
                    cmd.Parameters.AddWithValue("@VendorPN", productionItem.VendorPN);
                    cmd.Parameters.AddWithValue("@NewShortDescription", productionItem.NewShortDescription);
                    cmd.Parameters.AddWithValue("@NewLongDescription", productionItem.NewLongDescription);
                    cmd.Parameters.AddWithValue("@MissingWords", productionItem.MissingWords);
                    cmd.Parameters.AddWithValue("@Noun", productionItem.Noun);
                    cmd.Parameters.AddWithValue("@Modifier", productionItem.Modifier);
                    cmd.Parameters.AddWithValue("@Status", productionItem.Status);
                    cmd.Parameters.AddWithValue("@Level", productionItem.Level);
                    cmd.Parameters.AddWithValue("@MFRName1", productionItem.MFRName1);
                    cmd.Parameters.AddWithValue("@MFRPN1", productionItem.MFRPN1);
                    cmd.Parameters.AddWithValue("@MFRName2", productionItem.MFRName2);
                    cmd.Parameters.AddWithValue("@MFRPN2", productionItem.MFRPN2);
                    cmd.Parameters.AddWithValue("@MFRName3", productionItem.MFRName3);
                    cmd.Parameters.AddWithValue("@MFRPN3", productionItem.MFRPN3);
                    cmd.Parameters.AddWithValue("@VendorName1", productionItem.VendorName1);
                    cmd.Parameters.AddWithValue("@VendorPN1", productionItem.VendorPN1);
                    cmd.Parameters.AddWithValue("@VendorName2", productionItem.VendorName2);
                    cmd.Parameters.AddWithValue("@VendorPN2", productionItem.VendorPN2);
                    cmd.Parameters.AddWithValue("@VendorName3", productionItem.VendorName3);
                    cmd.Parameters.AddWithValue("@VendorPN3", productionItem.VendorPN3);
                    cmd.Parameters.AddWithValue("@AdditionalInfo", productionItem.AdditionalInfo);
                    cmd.Parameters.AddWithValue("@AdditionalInfoFromWeb", productionItem.AdditionalInfoFromWeb);
                    cmd.Parameters.AddWithValue("@UNSPSCCode", productionItem.UNSPSCCode);
                    cmd.Parameters.AddWithValue("@UNSPSCCategory", productionItem.UNSPSCCategory);
                    cmd.Parameters.AddWithValue("@WebRefURL1", productionItem.WebRefURL1);
                    cmd.Parameters.AddWithValue("@WebRefURL2", productionItem.WebRefURL2);
                    cmd.Parameters.AddWithValue("@WebRefURL3", productionItem.WebRefURL3);
                    cmd.Parameters.AddWithValue("@PDFURL", productionItem.PDFURL);
                    cmd.Parameters.AddWithValue("@Remarks", productionItem.Remarks);
                    cmd.Parameters.AddWithValue("@Query", productionItem.Query);
                    cmd.Parameters.AddWithValue("@Application", productionItem.Application);
                    cmd.Parameters.AddWithValue("@DWG", productionItem.DWG);
                    cmd.Parameters.AddWithValue("@ItemNo", productionItem.ItemNo);
                    cmd.Parameters.AddWithValue("@OtherNo", productionItem.OtherNo);
                    cmd.Parameters.AddWithValue("@POS", productionItem.POS);
                    cmd.Parameters.AddWithValue("@SerialNo", productionItem.SerialNo);
                    cmd.Parameters.AddWithValue("@KKSCode", productionItem.KKSCode);
                    cmd.Parameters.AddWithValue("@AssemblyOrPart", productionItem.AssemblyOrPart);
                    cmd.Parameters.AddWithValue("@BOM", productionItem.BOM);
                    cmd.Parameters.AddWithValue("@GreenItems", productionItem.GreenItems);
                    cmd.Parameters.Add(new SqlParameter("@AttributeNameAndValue", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", productionItem.UserID);
                    #endregion

                    //Calling sp to update project column read only status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
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

        #region Production Item Details
        [HttpGet]
        [Route("ProductionItemDetails/{ProductionItemID}")]
        public IHttpActionResult ProductionItemDetails(long ProductionItemID)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Item Attributes
                List<ItemAttribute> ItemAttributesList = new List<ItemAttribute>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    #region Item Attributes
                    //Initialize command object
                    SqlCommand cmdItemAttributes = new SqlCommand();
                    cmdItemAttributes.Connection = conn;
                    cmdItemAttributes.CommandType = CommandType.StoredProcedure;
                    cmdItemAttributes.CommandText = "spProductionItemAttributeDetails";

                    #region Adding Stored Procedure Parameters
                    cmdItemAttributes.Parameters.AddWithValue("@ProductionItemID", ProductionItemID);
                    #endregion

                    //Call sp to get all Item Attributes
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmdItemAttributes.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ItemAttribute ItemAttribute = new ItemAttribute();
                        ItemAttribute.AttributeName = sqlReader["AttributeName"].ToString();
                        ItemAttribute.AttributeValue = sqlReader["AttributeValue"].ToString();
                        ItemAttribute.QCAttributeValue = sqlReader["QCAttributeValue"].ToString();
                        ItemAttribute.QCAttributeValueComments = sqlReader["QCAttributeValueComments"].ToString();
                        ItemAttributesList.Add(ItemAttribute);
                    }
                    conn.Close();
                    #endregion

                    #region Production Item
                    ProductionItem productionItem = new ProductionItem();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionItemDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@ProductionItemID", ProductionItemID);

                    //Call sp to get all project details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        productionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        productionItem.CustomerCode = sqlReader["CustomerCode"].ToString();
                        productionItem.ProjectCode = sqlReader["ProjectCode"].ToString();
                        productionItem.BatchNo = sqlReader["BatchNo"].ToString();
                        productionItem.ProductionUser = sqlReader["ProductionUser"].ToString();
                        productionItem.PreviousProductionItemID = sqlReader["PreviousProductionItemID"] == DBNull.Value ? null : (Int64?)Convert.ToInt64(sqlReader["PreviousProductionItemID"]);
                        productionItem.NextProductionItemID = sqlReader["NextProductionItemID"] == DBNull.Value ? null : (Int64?)Convert.ToInt64(sqlReader["NextProductionItemID"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.CustomColumnName1 = sqlReader["CustomColumnName1"].ToString();
                        productionItem.CustomColumnName1Value = sqlReader["CustomColumnName1Value"].ToString();
                        productionItem.CustomColumnName2 = sqlReader["CustomColumnName2"].ToString();
                        productionItem.CustomColumnName2Value = sqlReader["CustomColumnName2Value"].ToString();
                        productionItem.CustomColumnName3 = sqlReader["CustomColumnName3"].ToString();
                        productionItem.CustomColumnName3Value = sqlReader["CustomColumnName3Value"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Query = sqlReader["Query"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();
                        productionItem.UNSPSCVersion = sqlReader["UNSPSCVersion"].ToString();
                        productionItem.ItemAttributes = ItemAttributesList;
                        conn.Close();
                    }

                    //return Production to the request
                    return Ok(productionItem);

                    #endregion
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

        #region Noun, Modifier List
        [HttpGet]
        [Route("NounModifierList")]
        public IHttpActionResult NounModifierList(string CustomerCode, string ProjectCode)
        {
            try
            {
                List<NounModifier> NounModifierList = new List<NounModifier>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spNounModifiersOfProject";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);

                    //Calling sp to get list of noun, modifier
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounModifier nounModifier = new NounModifier();
                        nounModifier.Noun = sqlReader["Noun"].ToString();
                        nounModifier.Modifier = sqlReader["Modifier"].ToString();
                        NounModifierList.Add(nounModifier);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(NounModifierList);
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

        #region Noun, Modifier Attribute List
        [HttpGet]
        [Route("NounModifierAttributeList")]
        public IHttpActionResult NounModifierAttributeList(string CustomerCode, string ProjectCode, string Noun, string Modifier)
        {
            try
            {
                List<NounModifierAttribute> NounModifierAttributeList = new List<NounModifierAttribute>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spNounModifierAttributeOfProject";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);

                    //Calling sp to get list of noun, modifier attribute list
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounModifierAttribute nounModifierAttribute = new NounModifierAttribute();
                        nounModifierAttribute.AttributeName = sqlReader["Attribute"].ToString();
                        NounModifierAttributeList.Add(nounModifierAttribute);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(NounModifierAttributeList);
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

        #region Read All Items of Production User from the Project Allocation of Selected Project
        [HttpGet]
        [Route("ProductionItemDetailsOfProductionUserFromAllocation/{CustomerCode}/{ProjectCode}/{ProductionUser}/{BatchNo?}")]
        public IHttpActionResult ProductionItemDetailsOfProductionUserFromAllocation(string CustomerCode, string ProjectCode, string ProductionUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Production Item List
                List<ProductionItem> ProductionItemList = new List<ProductionItem>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersAllocationProductionDataTemplate";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionItem ProductionItem = new ProductionItem();

                        ProductionItem.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        ProductionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        ProductionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        ProductionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        ProductionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        ProductionItem.UOM = sqlReader["UOM"].ToString();
                        ProductionItem.MFRName = sqlReader["MFRName"].ToString();
                        ProductionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        ProductionItem.VendorName = sqlReader["VendorName"].ToString();
                        ProductionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        ProductionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        ProductionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        ProductionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        ProductionItem.Noun = sqlReader["Noun"].ToString();
                        ProductionItem.Modifier = sqlReader["Modifier"].ToString();
                        ProductionItem.Status = sqlReader["Status"].ToString();
                        ProductionItem.Level = sqlReader["Level"].ToString();
                        ProductionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        ProductionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        ProductionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        ProductionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        ProductionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        ProductionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        ProductionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        ProductionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        ProductionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        ProductionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        ProductionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        ProductionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        ProductionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        ProductionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        ProductionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        ProductionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        ProductionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        ProductionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        ProductionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        ProductionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        ProductionItem.Remarks = sqlReader["Remarks"].ToString();
                        ProductionItem.Query = sqlReader["Query"].ToString();
                        ProductionItem.AttributeName1 = sqlReader["AttributeName1"].ToString();
                        ProductionItem.AttributeValue1 = sqlReader["AttributeValue1"].ToString();
                        ProductionItem.AttributeName2 = sqlReader["AttributeName2"].ToString();
                        ProductionItem.AttributeValue2 = sqlReader["AttributeValue2"].ToString();
                        ProductionItem.AttributeName3 = sqlReader["AttributeName3"].ToString();
                        ProductionItem.AttributeValue3 = sqlReader["AttributeValue3"].ToString();
                        ProductionItem.AttributeName4 = sqlReader["AttributeName4"].ToString();
                        ProductionItem.AttributeValue4 = sqlReader["AttributeValue4"].ToString();
                        ProductionItem.AttributeName5 = sqlReader["AttributeName5"].ToString();
                        ProductionItem.AttributeValue5 = sqlReader["AttributeValue5"].ToString();
                        ProductionItem.AttributeName6 = sqlReader["AttributeName6"].ToString();
                        ProductionItem.AttributeValue6 = sqlReader["AttributeValue6"].ToString();
                        ProductionItem.AttributeName7 = sqlReader["AttributeName7"].ToString();
                        ProductionItem.AttributeValue7 = sqlReader["AttributeValue7"].ToString();
                        ProductionItem.AttributeName8 = sqlReader["AttributeName8"].ToString();
                        ProductionItem.AttributeValue8 = sqlReader["AttributeValue8"].ToString();
                        ProductionItem.AttributeName9 = sqlReader["AttributeName9"].ToString();
                        ProductionItem.AttributeValue9 = sqlReader["AttributeValue9"].ToString();
                        ProductionItem.AttributeName10 = sqlReader["AttributeName10"].ToString();
                        ProductionItem.AttributeValue10 = sqlReader["AttributeValue10"].ToString();
                        ProductionItem.AttributeName11 = sqlReader["AttributeName11"].ToString();
                        ProductionItem.AttributeValue11 = sqlReader["AttributeValue11"].ToString();
                        ProductionItem.AttributeName12 = sqlReader["AttributeName12"].ToString();
                        ProductionItem.AttributeValue12 = sqlReader["AttributeValue12"].ToString();
                        ProductionItem.AttributeName13 = sqlReader["AttributeName13"].ToString();
                        ProductionItem.AttributeValue13 = sqlReader["AttributeValue13"].ToString();
                        ProductionItem.AttributeName14 = sqlReader["AttributeName14"].ToString();
                        ProductionItem.AttributeValue14 = sqlReader["AttributeValue14"].ToString();
                        ProductionItem.AttributeName15 = sqlReader["AttributeName15"].ToString();
                        ProductionItem.AttributeValue15 = sqlReader["AttributeValue15"].ToString();
                        ProductionItem.AttributeName16 = sqlReader["AttributeName16"].ToString();
                        ProductionItem.AttributeValue16 = sqlReader["AttributeValue16"].ToString();
                        ProductionItem.AttributeName17 = sqlReader["AttributeName17"].ToString();
                        ProductionItem.AttributeValue17 = sqlReader["AttributeValue17"].ToString();
                        ProductionItem.AttributeName18 = sqlReader["AttributeName18"].ToString();
                        ProductionItem.AttributeValue18 = sqlReader["AttributeValue18"].ToString();
                        ProductionItem.AttributeName19 = sqlReader["AttributeName19"].ToString();
                        ProductionItem.AttributeValue19 = sqlReader["AttributeValue19"].ToString();
                        ProductionItem.AttributeName20 = sqlReader["AttributeName20"].ToString();
                        ProductionItem.AttributeValue20 = sqlReader["AttributeValue20"].ToString();
                        ProductionItem.AttributeName21 = sqlReader["AttributeName21"].ToString();
                        ProductionItem.AttributeValue21 = sqlReader["AttributeValue21"].ToString();
                        ProductionItem.AttributeName22 = sqlReader["AttributeName22"].ToString();
                        ProductionItem.AttributeValue22 = sqlReader["AttributeValue22"].ToString();
                        ProductionItem.AttributeName23 = sqlReader["AttributeName23"].ToString();
                        ProductionItem.AttributeValue23 = sqlReader["AttributeValue23"].ToString();
                        ProductionItem.AttributeName24 = sqlReader["AttributeName24"].ToString();
                        ProductionItem.AttributeValue24 = sqlReader["AttributeValue24"].ToString();
                        ProductionItem.AttributeName25 = sqlReader["AttributeName25"].ToString();
                        ProductionItem.AttributeValue25 = sqlReader["AttributeValue25"].ToString();
                        ProductionItem.AttributeName26 = sqlReader["AttributeName26"].ToString();
                        ProductionItem.AttributeValue26 = sqlReader["AttributeValue26"].ToString();
                        ProductionItem.AttributeName27 = sqlReader["AttributeName27"].ToString();
                        ProductionItem.AttributeValue27 = sqlReader["AttributeValue27"].ToString();
                        ProductionItem.AttributeName28 = sqlReader["AttributeName28"].ToString();
                        ProductionItem.AttributeValue28 = sqlReader["AttributeValue28"].ToString();
                        ProductionItem.AttributeName29 = sqlReader["AttributeName29"].ToString();
                        ProductionItem.AttributeValue29 = sqlReader["AttributeValue29"].ToString();
                        ProductionItem.AttributeName30 = sqlReader["AttributeName30"].ToString();
                        ProductionItem.AttributeValue30 = sqlReader["AttributeValue30"].ToString();
                        ProductionItem.AttributeName31 = sqlReader["AttributeName31"].ToString();
                        ProductionItem.AttributeValue31 = sqlReader["AttributeValue31"].ToString();
                        ProductionItem.AttributeName32 = sqlReader["AttributeName32"].ToString();
                        ProductionItem.AttributeValue32 = sqlReader["AttributeValue32"].ToString();
                        ProductionItem.AttributeName33 = sqlReader["AttributeName33"].ToString();
                        ProductionItem.AttributeValue33 = sqlReader["AttributeValue33"].ToString();
                        ProductionItem.AttributeName34 = sqlReader["AttributeName34"].ToString();
                        ProductionItem.AttributeValue34 = sqlReader["AttributeValue34"].ToString();
                        ProductionItem.AttributeName35 = sqlReader["AttributeName35"].ToString();
                        ProductionItem.AttributeValue35 = sqlReader["AttributeValue35"].ToString();
                        ProductionItem.AttributeName36 = sqlReader["AttributeName36"].ToString();
                        ProductionItem.AttributeValue36 = sqlReader["AttributeValue36"].ToString();
                        ProductionItem.AttributeName37 = sqlReader["AttributeName37"].ToString();
                        ProductionItem.AttributeValue37 = sqlReader["AttributeValue37"].ToString();
                        ProductionItem.AttributeName38 = sqlReader["AttributeName38"].ToString();
                        ProductionItem.AttributeValue38 = sqlReader["AttributeValue38"].ToString();
                        ProductionItem.AttributeName39 = sqlReader["AttributeName39"].ToString();
                        ProductionItem.AttributeValue39 = sqlReader["AttributeValue39"].ToString();
                        ProductionItem.AttributeName40 = sqlReader["AttributeName40"].ToString();
                        ProductionItem.AttributeValue40 = sqlReader["AttributeValue40"].ToString();
                        ProductionItem.AttributeName41 = sqlReader["AttributeName41"].ToString();
                        ProductionItem.AttributeValue41 = sqlReader["AttributeValue41"].ToString();
                        ProductionItem.AttributeName42 = sqlReader["AttributeName42"].ToString();
                        ProductionItem.AttributeValue42 = sqlReader["AttributeValue42"].ToString();
                        ProductionItem.AttributeName43 = sqlReader["AttributeName43"].ToString();
                        ProductionItem.AttributeValue43 = sqlReader["AttributeValue43"].ToString();
                        ProductionItem.AttributeName44 = sqlReader["AttributeName44"].ToString();
                        ProductionItem.AttributeValue44 = sqlReader["AttributeValue44"].ToString();
                        ProductionItem.AttributeName45 = sqlReader["AttributeName45"].ToString();
                        ProductionItem.AttributeValue45 = sqlReader["AttributeValue45"].ToString();
                        ProductionItem.AttributeName46 = sqlReader["AttributeName46"].ToString();
                        ProductionItem.AttributeValue46 = sqlReader["AttributeValue46"].ToString();
                        ProductionItem.AttributeName47 = sqlReader["AttributeName47"].ToString();
                        ProductionItem.AttributeValue47 = sqlReader["AttributeValue47"].ToString();
                        ProductionItem.AttributeName48 = sqlReader["AttributeName48"].ToString();
                        ProductionItem.AttributeValue48 = sqlReader["AttributeValue48"].ToString();
                        ProductionItem.AttributeName49 = sqlReader["AttributeName49"].ToString();
                        ProductionItem.AttributeValue49 = sqlReader["AttributeValue49"].ToString();
                        ProductionItem.AttributeName50 = sqlReader["AttributeName50"].ToString();
                        ProductionItem.AttributeValue50 = sqlReader["AttributeValue50"].ToString();

                        ProductionItemList.Add(ProductionItem);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProductionItemList);

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

        #region Read Production User's Project Production Data
        [HttpGet]
        [Route("ProductionUsersProjectProductionData")]
        public IHttpActionResult ProductionUsersProjectProductionData(string CustomerCode, string ProjectCode, long ProductionAllocationID, string ProductionUser, string BatchNo = "", int PageNo = 1, int PageSize = 20, string Status = "A")          //I-InProcess, C-Completed, Q-Query, M-Moved To QC, A-All
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Production Item List
                List<ProductionItem> ProductionItemList = new List<ProductionItem>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersProjectProductionData";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);
                    cmd.Parameters.AddWithValue("@PageNo", PageNo);
                    cmd.Parameters.AddWithValue("@PageSize", PageSize);
                    cmd.Parameters.AddWithValue("@IsToFetchAttributeDetails", 0);
                    cmd.Parameters.AddWithValue("@Status", Status);

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem productionItem = new ProductionItem();

                        productionItem.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.CustomColumnName1 = sqlReader["CustomColumnName1"].ToString();
                        productionItem.CustomColumnName1Value = sqlReader["CustomColumnName1Value"].ToString();
                        productionItem.CustomColumnName2 = sqlReader["CustomColumnName2"].ToString();
                        productionItem.CustomColumnName2Value = sqlReader["CustomColumnName2Value"].ToString();
                        productionItem.CustomColumnName3 = sqlReader["CustomColumnName3"].ToString();
                        productionItem.CustomColumnName3Value = sqlReader["CustomColumnName3Value"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.IsMovedToQC = sqlReader["IsMovedToQC"].ToString();
                        productionItem.TotalRowsCount = Convert.ToInt32(sqlReader["TotalCount"]);       //no need of this in excel export code

                        ProductionItemList.Add(productionItem);
                        #endregion
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProductionItemList);

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

        #region Download Project Production Item Details of User
        [HttpGet]
        [Route("DownloadProjectProductionItemDetailsOfUser")]
        public HttpResponseMessage DownloadProjectProductionItemDetailsOfUser(string CustomerCode, string ProjectCode, long ProductionAllocationID, string ProductionUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                string FileName = string.Empty;

                if (string.IsNullOrEmpty(BatchNo))
                    FileName = "ProductionItemDetails_" + CustomerCode + '_' + ProjectCode + '_' + ProductionUser + ".xlsx";
                else
                    FileName = "ProductionItemDetails_" + CustomerCode + '_' + ProjectCode + '_' + BatchNo + '_' + ProductionUser + ".xlsx";

                //Create a list to hold Production Item List
                List<ProductionItem> ProductionItemList = new List<ProductionItem>();

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                #region Setting Styles
                Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleInputColumnsHeader = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleLeftAlignDataWithoutBorder = ws.Cells[0, 0].GetStyle();
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

                styleInputColumnsHeader.IsTextWrapped = true;
                styleInputColumnsHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleInputColumnsHeader.VerticalAlignment = TextAlignmentType.Center;
                styleInputColumnsHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 211, 211, 211);
                styleInputColumnsHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 211, 211, 211);
                styleInputColumnsHeader.Pattern = BackgroundType.VerticalStripe;
                styleInputColumnsHeader.Font.Color = System.Drawing.Color.Black;
                styleInputColumnsHeader.Font.IsBold = true;
                styleInputColumnsHeader.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleInputColumnsHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleInputColumnsHeader.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleInputColumnsHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleInputColumnsHeader.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleInputColumnsHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleInputColumnsHeader.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleInputColumnsHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                styleCenterAlignData.HorizontalAlignment = TextAlignmentType.Center;
                styleCenterAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleCenterAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleCenterAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                styleLeftAlignDataWithoutBorder.HorizontalAlignment = TextAlignmentType.Left;

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
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersProjectProductionData";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);
                    cmd.Parameters.AddWithValue("@PageNo", 1);
                    cmd.Parameters.AddWithValue("@PageSize", 100000);
                    cmd.Parameters.AddWithValue("@IsToFetchAttributeDetails", 1);
                    cmd.Parameters.AddWithValue("@Status", "A");

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem productionItem = new ProductionItem();

                        productionItem.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();
                        productionItem.AttributeName1 = sqlReader["AttributeName1"].ToString();
                        productionItem.AttributeValue1 = sqlReader["AttributeValue1"].ToString();
                        productionItem.AttributeName2 = sqlReader["AttributeName2"].ToString();
                        productionItem.AttributeValue2 = sqlReader["AttributeValue2"].ToString();
                        productionItem.AttributeName3 = sqlReader["AttributeName3"].ToString();
                        productionItem.AttributeValue3 = sqlReader["AttributeValue3"].ToString();
                        productionItem.AttributeName4 = sqlReader["AttributeName4"].ToString();
                        productionItem.AttributeValue4 = sqlReader["AttributeValue4"].ToString();
                        productionItem.AttributeName5 = sqlReader["AttributeName5"].ToString();
                        productionItem.AttributeValue5 = sqlReader["AttributeValue5"].ToString();
                        productionItem.AttributeName6 = sqlReader["AttributeName6"].ToString();
                        productionItem.AttributeValue6 = sqlReader["AttributeValue6"].ToString();
                        productionItem.AttributeName7 = sqlReader["AttributeName7"].ToString();
                        productionItem.AttributeValue7 = sqlReader["AttributeValue7"].ToString();
                        productionItem.AttributeName8 = sqlReader["AttributeName8"].ToString();
                        productionItem.AttributeValue8 = sqlReader["AttributeValue8"].ToString();
                        productionItem.AttributeName9 = sqlReader["AttributeName9"].ToString();
                        productionItem.AttributeValue9 = sqlReader["AttributeValue9"].ToString();
                        productionItem.AttributeName10 = sqlReader["AttributeName10"].ToString();
                        productionItem.AttributeValue10 = sqlReader["AttributeValue10"].ToString();
                        productionItem.AttributeName11 = sqlReader["AttributeName11"].ToString();
                        productionItem.AttributeValue11 = sqlReader["AttributeValue11"].ToString();
                        productionItem.AttributeName12 = sqlReader["AttributeName12"].ToString();
                        productionItem.AttributeValue12 = sqlReader["AttributeValue12"].ToString();
                        productionItem.AttributeName13 = sqlReader["AttributeName13"].ToString();
                        productionItem.AttributeValue13 = sqlReader["AttributeValue13"].ToString();
                        productionItem.AttributeName14 = sqlReader["AttributeName14"].ToString();
                        productionItem.AttributeValue14 = sqlReader["AttributeValue14"].ToString();
                        productionItem.AttributeName15 = sqlReader["AttributeName15"].ToString();
                        productionItem.AttributeValue15 = sqlReader["AttributeValue15"].ToString();
                        productionItem.AttributeName16 = sqlReader["AttributeName16"].ToString();
                        productionItem.AttributeValue16 = sqlReader["AttributeValue16"].ToString();
                        productionItem.AttributeName17 = sqlReader["AttributeName17"].ToString();
                        productionItem.AttributeValue17 = sqlReader["AttributeValue17"].ToString();
                        productionItem.AttributeName18 = sqlReader["AttributeName18"].ToString();
                        productionItem.AttributeValue18 = sqlReader["AttributeValue18"].ToString();
                        productionItem.AttributeName19 = sqlReader["AttributeName19"].ToString();
                        productionItem.AttributeValue19 = sqlReader["AttributeValue19"].ToString();
                        productionItem.AttributeName20 = sqlReader["AttributeName20"].ToString();
                        productionItem.AttributeValue20 = sqlReader["AttributeValue20"].ToString();
                        productionItem.AttributeName21 = sqlReader["AttributeName21"].ToString();
                        productionItem.AttributeValue21 = sqlReader["AttributeValue21"].ToString();
                        productionItem.AttributeName22 = sqlReader["AttributeName22"].ToString();
                        productionItem.AttributeValue22 = sqlReader["AttributeValue22"].ToString();
                        productionItem.AttributeName23 = sqlReader["AttributeName23"].ToString();
                        productionItem.AttributeValue23 = sqlReader["AttributeValue23"].ToString();
                        productionItem.AttributeName24 = sqlReader["AttributeName24"].ToString();
                        productionItem.AttributeValue24 = sqlReader["AttributeValue24"].ToString();
                        productionItem.AttributeName25 = sqlReader["AttributeName25"].ToString();
                        productionItem.AttributeValue25 = sqlReader["AttributeValue25"].ToString();
                        productionItem.AttributeName26 = sqlReader["AttributeName26"].ToString();
                        productionItem.AttributeValue26 = sqlReader["AttributeValue26"].ToString();
                        productionItem.AttributeName27 = sqlReader["AttributeName27"].ToString();
                        productionItem.AttributeValue27 = sqlReader["AttributeValue27"].ToString();
                        productionItem.AttributeName28 = sqlReader["AttributeName28"].ToString();
                        productionItem.AttributeValue28 = sqlReader["AttributeValue28"].ToString();
                        productionItem.AttributeName29 = sqlReader["AttributeName29"].ToString();
                        productionItem.AttributeValue29 = sqlReader["AttributeValue29"].ToString();
                        productionItem.AttributeName30 = sqlReader["AttributeName30"].ToString();
                        productionItem.AttributeValue30 = sqlReader["AttributeValue30"].ToString();
                        productionItem.AttributeName31 = sqlReader["AttributeName31"].ToString();
                        productionItem.AttributeValue31 = sqlReader["AttributeValue31"].ToString();
                        productionItem.AttributeName32 = sqlReader["AttributeName32"].ToString();
                        productionItem.AttributeValue32 = sqlReader["AttributeValue32"].ToString();
                        productionItem.AttributeName33 = sqlReader["AttributeName33"].ToString();
                        productionItem.AttributeValue33 = sqlReader["AttributeValue33"].ToString();
                        productionItem.AttributeName34 = sqlReader["AttributeName34"].ToString();
                        productionItem.AttributeValue34 = sqlReader["AttributeValue34"].ToString();
                        productionItem.AttributeName35 = sqlReader["AttributeName35"].ToString();
                        productionItem.AttributeValue35 = sqlReader["AttributeValue35"].ToString();
                        productionItem.AttributeName36 = sqlReader["AttributeName36"].ToString();
                        productionItem.AttributeValue36 = sqlReader["AttributeValue36"].ToString();
                        productionItem.AttributeName37 = sqlReader["AttributeName37"].ToString();
                        productionItem.AttributeValue37 = sqlReader["AttributeValue37"].ToString();
                        productionItem.AttributeName38 = sqlReader["AttributeName38"].ToString();
                        productionItem.AttributeValue38 = sqlReader["AttributeValue38"].ToString();
                        productionItem.AttributeName39 = sqlReader["AttributeName39"].ToString();
                        productionItem.AttributeValue39 = sqlReader["AttributeValue39"].ToString();
                        productionItem.AttributeName40 = sqlReader["AttributeName40"].ToString();
                        productionItem.AttributeValue40 = sqlReader["AttributeValue40"].ToString();
                        productionItem.AttributeName41 = sqlReader["AttributeName41"].ToString();
                        productionItem.AttributeValue41 = sqlReader["AttributeValue41"].ToString();
                        productionItem.AttributeName42 = sqlReader["AttributeName42"].ToString();
                        productionItem.AttributeValue42 = sqlReader["AttributeValue42"].ToString();
                        productionItem.AttributeName43 = sqlReader["AttributeName43"].ToString();
                        productionItem.AttributeValue43 = sqlReader["AttributeValue43"].ToString();
                        productionItem.AttributeName44 = sqlReader["AttributeName44"].ToString();
                        productionItem.AttributeValue44 = sqlReader["AttributeValue44"].ToString();
                        productionItem.AttributeName45 = sqlReader["AttributeName45"].ToString();
                        productionItem.AttributeValue45 = sqlReader["AttributeValue45"].ToString();
                        productionItem.AttributeName46 = sqlReader["AttributeName46"].ToString();
                        productionItem.AttributeValue46 = sqlReader["AttributeValue46"].ToString();
                        productionItem.AttributeName47 = sqlReader["AttributeName47"].ToString();
                        productionItem.AttributeValue47 = sqlReader["AttributeValue47"].ToString();
                        productionItem.AttributeName48 = sqlReader["AttributeName48"].ToString();
                        productionItem.AttributeValue48 = sqlReader["AttributeValue48"].ToString();
                        productionItem.AttributeName49 = sqlReader["AttributeName49"].ToString();
                        productionItem.AttributeValue49 = sqlReader["AttributeValue49"].ToString();
                        productionItem.AttributeName50 = sqlReader["AttributeName50"].ToString();
                        productionItem.AttributeValue50 = sqlReader["AttributeValue50"].ToString();
                        productionItem.CustomColumnName1 = sqlReader["CustomColumnName1"].ToString();
                        productionItem.CustomColumnName1Value = sqlReader["CustomColumnName1Value"].ToString();
                        productionItem.CustomColumnName2 = sqlReader["CustomColumnName2"].ToString();
                        productionItem.CustomColumnName2Value = sqlReader["CustomColumnName2Value"].ToString();
                        productionItem.CustomColumnName3 = sqlReader["CustomColumnName3"].ToString();
                        productionItem.CustomColumnName3Value = sqlReader["CustomColumnName3Value"].ToString();

                        ProductionItemList.Add(productionItem);
                        #endregion
                    }
                    conn.Close();

                    #region Writing column headings
                    ws.Cells[0, 0].PutValue("Unique ID");
                    ws.Cells[0, 1].PutValue("Short Description");
                    ws.Cells[0, 2].PutValue("Long Description");
                    ws.Cells[0, 3].PutValue("UOM");
                    ws.Cells[0, 4].PutValue("MFR Name");
                    ws.Cells[0, 5].PutValue("MFR PN");
                    ws.Cells[0, 6].PutValue("Vendor Name");
                    ws.Cells[0, 7].PutValue("Vendor PN");
                    ws.Cells[0, 8].PutValue("New Short Description");
                    ws.Cells[0, 9].PutValue("New Long Description");
                    ws.Cells[0, 10].PutValue("Missing Words");
                    ws.Cells[0, 11].PutValue("Noun");
                    ws.Cells[0, 12].PutValue("Modifier");
                    ws.Cells[0, 13].PutValue("Status");
                    ws.Cells[0, 14].PutValue("Level");
                    ws.Cells[0, 15].PutValue("MFR Name1");
                    ws.Cells[0, 16].PutValue("MFR PN1");
                    ws.Cells[0, 17].PutValue("MFR Name2");
                    ws.Cells[0, 18].PutValue("MFR PN2");
                    ws.Cells[0, 19].PutValue("MFR Name3");
                    ws.Cells[0, 20].PutValue("MFR PN3");
                    ws.Cells[0, 21].PutValue("Vendor Name1");
                    ws.Cells[0, 22].PutValue("Vendor PN1");
                    ws.Cells[0, 23].PutValue("Vendor Name2");
                    ws.Cells[0, 24].PutValue("Vendor PN2");
                    ws.Cells[0, 25].PutValue("Vendor Name3");
                    ws.Cells[0, 26].PutValue("Vendor PN3");
                    ws.Cells[0, 27].PutValue("Additional Info");
                    ws.Cells[0, 28].PutValue("Additional Info From Web");
                    ws.Cells[0, 29].PutValue("UNSPSC Code");
                    ws.Cells[0, 30].PutValue("UNSPSC Category");
                    ws.Cells[0, 31].PutValue("Web Ref URL1");
                    ws.Cells[0, 32].PutValue("Web Ref URL2");
                    ws.Cells[0, 33].PutValue("Web Ref URL3");
                    ws.Cells[0, 34].PutValue("PDF URL");
                    ws.Cells[0, 35].PutValue("Remarks");
                    ws.Cells[0, 36].PutValue("Query");
                    ws.Cells[0, 37].PutValue("Application");
                    ws.Cells[0, 38].PutValue("DWG");
                    ws.Cells[0, 39].PutValue("POS");
                    ws.Cells[0, 40].PutValue("Item No.");
                    ws.Cells[0, 41].PutValue("Serial No.");
                    ws.Cells[0, 42].PutValue("Other No.");
                    ws.Cells[0, 43].PutValue("KKSCode");
                    ws.Cells[0, 44].PutValue("Assembly Or Part");
                    ws.Cells[0, 45].PutValue("BOM");
                    ws.Cells[0, 46].PutValue("Green Items");
                    int colNo = 47;
                    for (int attributeNo = 1; attributeNo <= 50; attributeNo++)
                    {
                        ws.Cells[0, colNo].PutValue("Attribute Name" + attributeNo.ToString());
                        ws.Cells[0, colNo + 1].PutValue("Attribute Value" + attributeNo.ToString());
                        colNo += 2;
                    }
                    if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName1))
                        ws.Cells[0, 147].PutValue(ProductionItemList.FirstOrDefault().CustomColumnName1);
                    if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName2))
                        ws.Cells[0, 148].PutValue(ProductionItemList.FirstOrDefault().CustomColumnName2);
                    if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName3))
                        ws.Cells[0, 149].PutValue(ProductionItemList.FirstOrDefault().CustomColumnName3);
                    for (int c = 0; c <= ws.Cells.MaxColumn; c++)
                    {
                        if (c == 0 || c == 1 || c == 2 || c == 4 || c == 5 || c == 6 || c == 7)       //color Input Columns to Grey color
                            ws.Cells[0, c].SetStyle(styleInputColumnsHeader);
                        else
                            ws.Cells[0, c].SetStyle(styleHeader);
                    }
                    #endregion

                    foreach (ProductionItem productionItem in ProductionItemList)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(productionItem.UniqueID);
                        ws.Cells[row, 1].PutValue(productionItem.ShortDescription);
                        ws.Cells[row, 2].PutValue(productionItem.LongDescription);
                        ws.Cells[row, 3].PutValue(productionItem.UOM);
                        ws.Cells[row, 4].PutValue(productionItem.MFRName);
                        ws.Cells[row, 5].PutValue(productionItem.MFRPN);
                        ws.Cells[row, 6].PutValue(productionItem.VendorName);
                        ws.Cells[row, 7].PutValue(productionItem.VendorPN);
                        ws.Cells[row, 8].PutValue(productionItem.NewShortDescription);
                        ws.Cells[row, 9].PutValue(productionItem.NewLongDescription);
                        ws.Cells[row, 10].PutValue(productionItem.MissingWords);
                        ws.Cells[row, 11].PutValue(productionItem.Noun);
                        ws.Cells[row, 12].PutValue(productionItem.Modifier);
                        ws.Cells[row, 13].PutValue(productionItem.Status);
                        ws.Cells[row, 14].PutValue(productionItem.Level);
                        ws.Cells[row, 15].PutValue(productionItem.MFRName1);
                        ws.Cells[row, 16].PutValue(productionItem.MFRPN1);
                        ws.Cells[row, 17].PutValue(productionItem.MFRName2);
                        ws.Cells[row, 18].PutValue(productionItem.MFRPN2);
                        ws.Cells[row, 19].PutValue(productionItem.MFRName3);
                        ws.Cells[row, 20].PutValue(productionItem.MFRPN3);
                        ws.Cells[row, 21].PutValue(productionItem.VendorName1);
                        ws.Cells[row, 22].PutValue(productionItem.VendorPN1);
                        ws.Cells[row, 23].PutValue(productionItem.VendorName2);
                        ws.Cells[row, 24].PutValue(productionItem.VendorPN2);
                        ws.Cells[row, 25].PutValue(productionItem.VendorName3);
                        ws.Cells[row, 26].PutValue(productionItem.VendorPN3);
                        ws.Cells[row, 27].PutValue(productionItem.AdditionalInfo);
                        ws.Cells[row, 28].PutValue(productionItem.AdditionalInfoFromWeb);
                        ws.Cells[row, 29].PutValue(productionItem.UNSPSCCode);
                        ws.Cells[row, 30].PutValue(productionItem.UNSPSCCategory);
                        ws.Cells[row, 31].PutValue(productionItem.WebRefURL1);
                        ws.Cells[row, 32].PutValue(productionItem.WebRefURL2);
                        ws.Cells[row, 33].PutValue(productionItem.WebRefURL3);
                        ws.Cells[row, 34].PutValue(productionItem.PDFURL);
                        ws.Cells[row, 35].PutValue(productionItem.Remarks);
                        ws.Cells[row, 36].PutValue(productionItem.Query);
                        ws.Cells[row, 37].PutValue(productionItem.Application);
                        ws.Cells[row, 38].PutValue(productionItem.DWG);
                        ws.Cells[row, 39].PutValue(productionItem.POS);
                        ws.Cells[row, 40].PutValue(productionItem.ItemNo);
                        ws.Cells[row, 41].PutValue(productionItem.SerialNo);
                        ws.Cells[row, 42].PutValue(productionItem.OtherNo);
                        ws.Cells[row, 43].PutValue(productionItem.KKSCode);
                        ws.Cells[row, 44].PutValue(productionItem.AssemblyOrPart);
                        ws.Cells[row, 45].PutValue(productionItem.BOM);
                        ws.Cells[row, 46].PutValue(productionItem.GreenItems);
                        ws.Cells[row, 47].PutValue(productionItem.AttributeName1);
                        ws.Cells[row, 48].PutValue(productionItem.AttributeValue1);
                        ws.Cells[row, 49].PutValue(productionItem.AttributeName2);
                        ws.Cells[row, 50].PutValue(productionItem.AttributeValue2);
                        ws.Cells[row, 51].PutValue(productionItem.AttributeName3);
                        ws.Cells[row, 52].PutValue(productionItem.AttributeValue3);
                        ws.Cells[row, 53].PutValue(productionItem.AttributeName4);
                        ws.Cells[row, 54].PutValue(productionItem.AttributeValue4);
                        ws.Cells[row, 55].PutValue(productionItem.AttributeName5);
                        ws.Cells[row, 56].PutValue(productionItem.AttributeValue5);
                        ws.Cells[row, 57].PutValue(productionItem.AttributeName6);
                        ws.Cells[row, 58].PutValue(productionItem.AttributeValue6);
                        ws.Cells[row, 59].PutValue(productionItem.AttributeName7);
                        ws.Cells[row, 60].PutValue(productionItem.AttributeValue7);
                        ws.Cells[row, 61].PutValue(productionItem.AttributeName8);
                        ws.Cells[row, 62].PutValue(productionItem.AttributeValue8);
                        ws.Cells[row, 63].PutValue(productionItem.AttributeName9);
                        ws.Cells[row, 64].PutValue(productionItem.AttributeValue9);
                        ws.Cells[row, 65].PutValue(productionItem.AttributeName10);
                        ws.Cells[row, 66].PutValue(productionItem.AttributeValue10);
                        ws.Cells[row, 67].PutValue(productionItem.AttributeName11);
                        ws.Cells[row, 68].PutValue(productionItem.AttributeValue11);
                        ws.Cells[row, 69].PutValue(productionItem.AttributeName12);
                        ws.Cells[row, 70].PutValue(productionItem.AttributeValue12);
                        ws.Cells[row, 71].PutValue(productionItem.AttributeName13);
                        ws.Cells[row, 72].PutValue(productionItem.AttributeValue13);
                        ws.Cells[row, 73].PutValue(productionItem.AttributeName14);
                        ws.Cells[row, 74].PutValue(productionItem.AttributeValue14);
                        ws.Cells[row, 75].PutValue(productionItem.AttributeName15);
                        ws.Cells[row, 76].PutValue(productionItem.AttributeValue15);
                        ws.Cells[row, 77].PutValue(productionItem.AttributeName16);
                        ws.Cells[row, 78].PutValue(productionItem.AttributeValue16);
                        ws.Cells[row, 79].PutValue(productionItem.AttributeName17);
                        ws.Cells[row, 80].PutValue(productionItem.AttributeValue17);
                        ws.Cells[row, 81].PutValue(productionItem.AttributeName18);
                        ws.Cells[row, 82].PutValue(productionItem.AttributeValue18);
                        ws.Cells[row, 83].PutValue(productionItem.AttributeName19);
                        ws.Cells[row, 84].PutValue(productionItem.AttributeValue19);
                        ws.Cells[row, 85].PutValue(productionItem.AttributeName20);
                        ws.Cells[row, 86].PutValue(productionItem.AttributeValue20);
                        ws.Cells[row, 87].PutValue(productionItem.AttributeName21);
                        ws.Cells[row, 88].PutValue(productionItem.AttributeValue21);
                        ws.Cells[row, 89].PutValue(productionItem.AttributeName22);
                        ws.Cells[row, 90].PutValue(productionItem.AttributeValue22);
                        ws.Cells[row, 91].PutValue(productionItem.AttributeName23);
                        ws.Cells[row, 92].PutValue(productionItem.AttributeValue23);
                        ws.Cells[row, 93].PutValue(productionItem.AttributeName24);
                        ws.Cells[row, 94].PutValue(productionItem.AttributeValue24);
                        ws.Cells[row, 95].PutValue(productionItem.AttributeName25);
                        ws.Cells[row, 96].PutValue(productionItem.AttributeValue25);
                        ws.Cells[row, 97].PutValue(productionItem.AttributeName26);
                        ws.Cells[row, 98].PutValue(productionItem.AttributeValue26);
                        ws.Cells[row, 99].PutValue(productionItem.AttributeName27);
                        ws.Cells[row, 100].PutValue(productionItem.AttributeValue27);
                        ws.Cells[row, 101].PutValue(productionItem.AttributeName28);
                        ws.Cells[row, 102].PutValue(productionItem.AttributeValue28);
                        ws.Cells[row, 103].PutValue(productionItem.AttributeName29);
                        ws.Cells[row, 104].PutValue(productionItem.AttributeValue29);
                        ws.Cells[row, 105].PutValue(productionItem.AttributeName30);
                        ws.Cells[row, 106].PutValue(productionItem.AttributeValue30);
                        ws.Cells[row, 107].PutValue(productionItem.AttributeName31);
                        ws.Cells[row, 108].PutValue(productionItem.AttributeValue31);
                        ws.Cells[row, 109].PutValue(productionItem.AttributeName32);
                        ws.Cells[row, 110].PutValue(productionItem.AttributeValue32);
                        ws.Cells[row, 111].PutValue(productionItem.AttributeName33);
                        ws.Cells[row, 112].PutValue(productionItem.AttributeValue33);
                        ws.Cells[row, 113].PutValue(productionItem.AttributeName34);
                        ws.Cells[row, 114].PutValue(productionItem.AttributeValue34);
                        ws.Cells[row, 115].PutValue(productionItem.AttributeName35);
                        ws.Cells[row, 116].PutValue(productionItem.AttributeValue35);
                        ws.Cells[row, 117].PutValue(productionItem.AttributeName36);
                        ws.Cells[row, 118].PutValue(productionItem.AttributeValue36);
                        ws.Cells[row, 119].PutValue(productionItem.AttributeName37);
                        ws.Cells[row, 120].PutValue(productionItem.AttributeValue37);
                        ws.Cells[row, 121].PutValue(productionItem.AttributeName38);
                        ws.Cells[row, 122].PutValue(productionItem.AttributeValue38);
                        ws.Cells[row, 123].PutValue(productionItem.AttributeName39);
                        ws.Cells[row, 124].PutValue(productionItem.AttributeValue39);
                        ws.Cells[row, 125].PutValue(productionItem.AttributeName40);
                        ws.Cells[row, 126].PutValue(productionItem.AttributeValue40);
                        ws.Cells[row, 127].PutValue(productionItem.AttributeName41);
                        ws.Cells[row, 128].PutValue(productionItem.AttributeValue41);
                        ws.Cells[row, 129].PutValue(productionItem.AttributeName42);
                        ws.Cells[row, 130].PutValue(productionItem.AttributeValue42);
                        ws.Cells[row, 131].PutValue(productionItem.AttributeName43);
                        ws.Cells[row, 132].PutValue(productionItem.AttributeValue43);
                        ws.Cells[row, 133].PutValue(productionItem.AttributeName44);
                        ws.Cells[row, 134].PutValue(productionItem.AttributeValue44);
                        ws.Cells[row, 135].PutValue(productionItem.AttributeName45);
                        ws.Cells[row, 136].PutValue(productionItem.AttributeValue45);
                        ws.Cells[row, 137].PutValue(productionItem.AttributeName46);
                        ws.Cells[row, 138].PutValue(productionItem.AttributeValue46);
                        ws.Cells[row, 139].PutValue(productionItem.AttributeName47);
                        ws.Cells[row, 140].PutValue(productionItem.AttributeValue47);
                        ws.Cells[row, 141].PutValue(productionItem.AttributeName48);
                        ws.Cells[row, 142].PutValue(productionItem.AttributeValue48);
                        ws.Cells[row, 143].PutValue(productionItem.AttributeName49);
                        ws.Cells[row, 144].PutValue(productionItem.AttributeValue49);
                        ws.Cells[row, 145].PutValue(productionItem.AttributeName50);
                        ws.Cells[row, 146].PutValue(productionItem.AttributeValue50);
                        if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName1))
                            ws.Cells[row, 147].PutValue(productionItem.CustomColumnName1Value);
                        if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName2))
                            ws.Cells[row, 148].PutValue(productionItem.CustomColumnName2Value);
                        if (!string.IsNullOrEmpty(ProductionItemList.FirstOrDefault().CustomColumnName3))
                            ws.Cells[row, 149].PutValue(productionItem.CustomColumnName3Value);
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 8].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 9].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 10].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 12].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 13].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 25].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 27].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 28].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 29].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 31].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 32].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 33].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 35].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 36].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 37].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 38].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 39].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 40].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 41].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 42].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                        for (int col = 47; col <= ws.Cells.MaxColumn; col++)
                            ws.Cells[row, col].SetStyle(styleLeftAlignData);
                        #endregion

                        row++;
                    }
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
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Project MFR or Vendor Unique Names
        [HttpGet]
        [Route("ProjectMFRorVendorUniqueNames/{CustomerCode}/{ProjectCode}")]
        public IHttpActionResult ProjectMFRorVendorUniqueNames(string CustomerCode, string ProjectCode, string MFRorVendorFlag = "M")
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold MFR or Vendor Names List
                List<String> MFRorVendorNameList = new List<String>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectMFRVendorNameUniqueValues";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@MFRorVendorFlag", MFRorVendorFlag);

                    //Call sp to get the list of unique MFR or Vendor Names
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        MFRorVendorNameList.Add(sqlReader["MFRorVendorName"].ToString());
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(MFRorVendorNameList);
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

        #region Project Noun Modifier Unique Attribute Values
        [HttpPost]
        [Route("ProjectNounModifierUniqueAttributeValues")]
        public IHttpActionResult ProjectNounModifierUniqueAttributeValues(ProjectNounModifierAttribute projectNounModifierAttribute)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Project Noun Modifier Attribute Values List
                List<String> ProjectNounModifierUniqueAttributeValuesList = new List<String>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectNMsAttributeUniqueValues";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", projectNounModifierAttribute.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectNounModifierAttribute.ProjectCode);
                    cmd.Parameters.AddWithValue("@Noun", projectNounModifierAttribute.Noun);
                    cmd.Parameters.AddWithValue("@Modifier", projectNounModifierAttribute.Modifier);
                    cmd.Parameters.AddWithValue("@AttributeName", projectNounModifierAttribute.AttributeName);

                    //Call sp to get the list of unique Attribute values
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectNounModifierUniqueAttributeValuesList.Add(sqlReader["AttributeValue"].ToString());
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectNounModifierUniqueAttributeValuesList);
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

        #region Change Production Item Status
        [HttpPatch]
        [Route("ChangeProductionItemStatus/{ProductionItemID}/{Status}/{UserID}/{Level?}")]
        public HttpResponseMessage ChangeProductionItemStatus(long ProductionItemID, char Status, string UserID, string Level = "")
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionChangeItemStatus";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@ProductionItemID", ProductionItemID);
                    cmd.Parameters.AddWithValue("@Status", Status);
                    cmd.Parameters.AddWithValue("@Level", Level);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to update Production Item Status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
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

        #region Read Project Noun Modifier UNSPSC Code and Categories
        [HttpGet]
        [Route("ReadProjectNounModifierUNSPSCCodeAndCategories/{CustomerCode}/{ProjectCode}/{Noun}/{Modifier}")]
        public IHttpActionResult ReadProjectNounModifierUNSPSCCodeAndCategories(string CustomerCode, string ProjectCode, string Noun, string Modifier)
        {
            try
            {
                List<string> UNSPSCCodeCategoryList = new List<string>();
                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectNounModifiersUNSPSCCodeCategories";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);

                    //Calling sp to get list of UNSPSC Code and Categories
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        UNSPSCCodeCategoryList.Add(sqlReader["UNSPSCCodeCategory"].ToString());
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(UNSPSCCodeCategoryList.Distinct());
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

        #region Find duplicate unique id based on MFR/Vendor P/N
        [HttpPost]
        [Route("FindDuplicateUniqueID")]
        public HttpResponseMessage FindDuplicateUniqueID([FromBody] MFRVendorPNProjectModel model)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionFindDuplicateUniqueID";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.AddWithValue("@UniqueID", model.UniqueID);
                    cmd.Parameters.AddWithValue("@MFRPN1", model.MFRPN1);
                    cmd.Parameters.AddWithValue("@MFRPN2", model.MFRPN2);
                    cmd.Parameters.AddWithValue("@MFRPN3", model.MFRPN3);
                    cmd.Parameters.AddWithValue("@VendorPN1", model.VendorPN1);
                    cmd.Parameters.AddWithValue("@VendorPN2", model.VendorPN2);
                    cmd.Parameters.AddWithValue("@VendorPN3", model.VendorPN3);

                    //Calling sp to get the result
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.Trim() == string.Empty)
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
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

        #region Read Duplicate SKUs based on Selected Columns from selected project Production
        [HttpPost]
        [Route("FindDuplicatesOnSelectedColumns")]
        public IHttpActionResult FindDuplicatesOnSelectedColumns(DuplicateToFindOnColumnsModel model)
        {
            try
            {
                //Create a list to hold Production Item List
                List<ProductionItem> ProductionItemList = new List<ProductionItem>();
                System.Data.Common.DbDataReader sqlReader;

                #region Selected Columns
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<DuplicateToFindOnColumns>),
                                       new XmlRootAttribute("root"));

                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, model.ColumnNames);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionFindDuplicatesOnSelectedColumns";

                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.Add(new SqlParameter("@SelectedCoumnNames", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem ProductionItem = new ProductionItem();

                        ProductionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        ProductionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        ProductionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        ProductionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        ProductionItem.ProductionUser = sqlReader["ProductionUser"].ToString();
                        ProductionItem.UOM = sqlReader["UOM"].ToString();
                        ProductionItem.MFRName = sqlReader["MFRName"].ToString();
                        ProductionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        ProductionItem.VendorName = sqlReader["VendorName"].ToString();
                        ProductionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        ProductionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        ProductionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        ProductionItem.Noun = sqlReader["Noun"].ToString();
                        ProductionItem.Modifier = sqlReader["Modifier"].ToString();
                        ProductionItem.Status = sqlReader["Status"].ToString();
                        ProductionItem.Level = sqlReader["Level"].ToString();
                        ProductionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        ProductionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        ProductionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        ProductionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        ProductionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        ProductionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        ProductionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        ProductionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        ProductionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        ProductionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        ProductionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        ProductionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        ProductionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        ProductionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        ProductionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        ProductionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        ProductionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        ProductionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        ProductionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        ProductionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        ProductionItem.Remarks = sqlReader["Remarks"].ToString();

                        ProductionItemList.Add(ProductionItem);
                        #endregion
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProductionItemList);
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

        #region Read Duplicate SKUs based on Selected Columns from customer input file
        [HttpPost]
        [Route("ReadDuplicateSKUsBasedOnSelectedColumnsFromCustomerInputFile")]
        public IHttpActionResult ReadDuplicateSKUsBasedOnSelectedColumnsFromCustomerInputFile(DuplicateToFindOnColumnsModel model)
        {
            try
            {
                List<int> lstColumnNosOfSelectedColumns = new List<int>();
                int DuplicateRemarksColNo = -1, DuplicateRowSetCounter = 1, TotalDataRows = 0, DuplicateRowsCount = 0;
                bool cellValueMatched = false, AreAllColumnValuesBlank = false, IsDuplicateCounterPrinted = false;
                DataFormatConverter dataFormatConverter = new DataFormatConverter();

                #region Find the customer input file name
                string FileName = string.Empty;
                DirectoryInfo dirCustomerInputFile;

                if (string.IsNullOrEmpty(model.BatchNo))
                    FileName = model.CustomerCode + '_' + model.ProjectCode + "_CustomerInputFile.xlsx";
                else
                    FileName = model.CustomerCode + '_' + model.ProjectCode + '_' + model.BatchNo + "_CustomerInputFile.xlsx";

                if (string.IsNullOrEmpty(model.BatchNo))
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                else
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));

                string CustomerInputFilepath = dirCustomerInputFile + FileName;

                if (!File.Exists(CustomerInputFilepath))
                    return Content(HttpStatusCode.BadRequest, "Customer Input File Not Found");
                #endregion

                #region Copy the file to temp folder with unique name
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                string TempCustomerFileName = Guid.NewGuid().ToString() + ".xlsx";
                string TempCustomerFileFullPath = dirTemp.FullName + TempCustomerFileName;
                FileInfo fileInfo = new FileInfo(CustomerInputFilepath);
                fileInfo.CopyTo(TempCustomerFileFullPath);
                #endregion

                #region Open Temp Customer Input File and add Duplicate Remarks column
                Workbook wbTIF = new Workbook();             //Input File
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbTIF.LoadData(TempCustomerFileFullPath);
                var wsTIF = wbTIF.Worksheets[0];
                //wsTIF.Cells.DeleteBlankRows();
                TotalDataRows = wsTIF.Cells.MaxRow;

                DuplicateRemarksColNo = wsTIF.Cells.MaxColumn + 1;
                wsTIF.Cells[0, DuplicateRemarksColNo].PutValue("Duplicate Remarks");
                #endregion

                #region Find the Column Nos. of Selected Columns
                foreach (DuplicateToFindOnColumns columnName in model.ColumnNames)
                {
                    for (int col = 0; col <= wsTIF.Cells.MaxColumn; col++)
                    {
                        if (wsTIF.Cells[0, col].StringValue.Trim().ToUpper() == columnName.ColumnName.Trim().ToUpper())
                        {
                            lstColumnNosOfSelectedColumns.Add(col);
                            break;
                        }
                    }
                }
                #endregion

                #region Find the duplicate SKUs and write the row set no. in Duplicate Remarks column
                for (int mainRow = 1; mainRow <= wsTIF.Cells.MaxRow; mainRow++)
                {
                    if (IsDuplicateCounterPrinted)
                    {
                        DuplicateRowSetCounter++;
                        IsDuplicateCounterPrinted = false;
                    }

                    if (string.IsNullOrEmpty(wsTIF.Cells[mainRow, DuplicateRemarksColNo].StringValue))
                    {
                        AreAllColumnValuesBlank = false;
                        foreach (int c in lstColumnNosOfSelectedColumns)
                        {
                            if (!string.IsNullOrEmpty(wsTIF.Cells[mainRow, c].StringValue))
                                break;
                            else
                                AreAllColumnValuesBlank = true;
                        }

                        if (!AreAllColumnValuesBlank)
                        {
                            for (int iRow = mainRow + 1; iRow <= wsTIF.Cells.MaxRow; iRow++)
                            {
                                if (string.IsNullOrEmpty(wsTIF.Cells[iRow, DuplicateRemarksColNo].StringValue))
                                {
                                    cellValueMatched = false;
                                    foreach (int col in lstColumnNosOfSelectedColumns)
                                    {
                                        if (dataFormatConverter.RemoveSpecialCharacters(wsTIF.Cells[mainRow, col].StringValue.Trim().ToUpper()) == dataFormatConverter.RemoveSpecialCharacters(wsTIF.Cells[iRow, col].StringValue.Trim().ToUpper()))
                                            cellValueMatched = true;
                                        else
                                        {
                                            cellValueMatched = false;
                                            break;
                                        }
                                    }

                                    if (cellValueMatched)
                                    {
                                        if (string.IsNullOrEmpty(wsTIF.Cells[mainRow, DuplicateRemarksColNo].StringValue))
                                            wsTIF.Cells[mainRow, DuplicateRemarksColNo].PutValue("Duplicate Row Set " + DuplicateRowSetCounter.ToString());
                                        wsTIF.Cells[iRow, DuplicateRemarksColNo].PutValue("Duplicate Row Set " + DuplicateRowSetCounter.ToString());
                                        IsDuplicateCounterPrinted = true;
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Delete Duplicate Remarks blank rows from the temp Customer Input File
                for (int row = 1; row <= wsTIF.Cells.MaxRow;)
                {
                    if (string.IsNullOrEmpty(wsTIF.Cells[row, DuplicateRemarksColNo].StringValue.Trim()))
                    {
                        wsTIF.Cells.DeleteRow(row);
                        if (wsTIF.Cells.MaxRow == 0 || row == wsTIF.Cells.MaxRow)
                            break;
                    }
                    else
                        row++;
                }

                if (string.IsNullOrEmpty(wsTIF.Cells[wsTIF.Cells.MaxRow, DuplicateRemarksColNo].StringValue.Trim()))
                    wsTIF.Cells.DeleteRow(wsTIF.Cells.MaxRow);
                #endregion

                #region Write Duplicate Percentage
                DuplicateRowsCount = wsTIF.Cells.MaxRow;
                wsTIF.Cells[0, DuplicateRemarksColNo + 1].PutValue("Duplicate Percentage");
                wsTIF.Cells[1, DuplicateRemarksColNo + 1].PutValue((DuplicateRowsCount * 100.00) / TotalDataRows);
                #endregion

                #region Save the file changes and convert the file to datatable
                wbTIF.Save(TempCustomerFileFullPath);
                DataTable excelData = dataFormatConverter.ExcelToDataTable(TempCustomerFileFullPath);

                string ColumnsToBeSorted = string.Empty;
                foreach (DuplicateToFindOnColumns colName in model.ColumnNames)
                {
                    if (string.IsNullOrEmpty(ColumnsToBeSorted))
                        ColumnsToBeSorted = colName.ColumnName + " ASC";
                    else
                        ColumnsToBeSorted += "," + colName.ColumnName + " ASC";
                }
                #endregion

                DataView dv = excelData.DefaultView;
                dv.Sort = ColumnsToBeSorted;
                DataTable sortedDT = dv.ToTable();

                #region Delete the temp file
                FileInfo TempCustomerFileInfo = new FileInfo(TempCustomerFileFullPath);
                TempCustomerFileInfo.Delete();
                #endregion

                //return Ok(excelData);
                return Ok(sortedDT);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Move Selected SKUs to QC
        [HttpPatch]
        [Route("MoveSelectedSKUsToQC")]
        public HttpResponseMessage MoveSelectedSKUsToQC(MoveToQCModel model)
        {
            try
            {
                #region Production Item IDs
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ProductionItemIDModel>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, model.ProductionItemIDs);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion


                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionMoveToQC";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", model.ProductionAllocationID);
                    cmd.Parameters.Add(new SqlParameter("@ProductionItemIDs", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);
                    #endregion

                    //Calling sp to update Is Moved To QC flag to true
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
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

        #region Read all the SKUs Pending for QC from selected Project / Batch
        [HttpGet]
        [Route("ReadAllSKUsPendingForQCFromSelectedProjectOrBatch")]
        public IHttpActionResult ReadAllSKUsPendingForQCFromSelectedProjectOrBatch(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Production Item List
                List<ProductionItem> productionItemList = new List<ProductionItem>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersProjectProductionData";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@PageNo", 1);
                    cmd.Parameters.AddWithValue("@PageSize", 100000);
                    cmd.Parameters.AddWithValue("@Status", "M");

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem productionItem = new ProductionItem();

                        productionItem.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.ProductionUser = sqlReader["ProductionUser"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.IsMovedToQC = sqlReader["IsMovedToQC"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Query = sqlReader["Query"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();

                        productionItemList.Add(productionItem);
                        #endregion
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(productionItemList);
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

        #region Read Production Update List Search Fields
        [HttpGet]
        [Route("ReadProductionUpdateListSearchFields")]
        public IHttpActionResult ReadProductionUpdateListSearchFields()
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Search fields from Production Update List
                List<String> ProductionUpdateListSearchFields = new List<String>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUpdateListSearchFields";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Call sp to get the list of Production Update List Search Fields
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionUpdateListSearchFields.Add(sqlReader["SearchField"].ToString());
                    }

                    conn.Close();

                    //return list to the request
                    return Ok(ProductionUpdateListSearchFields);
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

        #region Read Production Update List by Searching Text in selected Search Field
        [HttpGet]
        [Route("SearchProductionUpdateList")]
        public IHttpActionResult SearchProductionUpdateList(string CustomerCode, string ProjectCode, long ProductionAllocationID, string ProductionUser, string SearchOn, string SearchText, string SortOn, string SortDirection = "ASC", string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Production Item List
                List<ProductionItem> ProductionItemList = new List<ProductionItem>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUpdateListSearchText";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser.Trim());
                    cmd.Parameters.AddWithValue("@SearchOn", SearchOn);
                    cmd.Parameters.AddWithValue("@SearchText", SearchText);
                    cmd.Parameters.AddWithValue("@SortOn", SortOn);
                    cmd.Parameters.AddWithValue("@SortDirection", SortDirection);

                    //Call sp to get all Production details of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem ProductionItem = new ProductionItem();

                        ProductionItem.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        ProductionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        ProductionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        ProductionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        ProductionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        ProductionItem.UOM = sqlReader["UOM"].ToString();
                        ProductionItem.MFRName = sqlReader["MFRName"].ToString();
                        ProductionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        ProductionItem.VendorName = sqlReader["VendorName"].ToString();
                        ProductionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        ProductionItem.CustomColumnName1 = sqlReader["CustomColumnName1"].ToString();
                        ProductionItem.CustomColumnName1Value = sqlReader["CustomColumnName1Value"].ToString();
                        ProductionItem.CustomColumnName2 = sqlReader["CustomColumnName2"].ToString();
                        ProductionItem.CustomColumnName2Value = sqlReader["CustomColumnName2Value"].ToString();
                        ProductionItem.CustomColumnName3 = sqlReader["CustomColumnName3"].ToString();
                        ProductionItem.CustomColumnName3Value = sqlReader["CustomColumnName3Value"].ToString();
                        ProductionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        ProductionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        ProductionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        ProductionItem.Noun = sqlReader["Noun"].ToString();
                        ProductionItem.Modifier = sqlReader["Modifier"].ToString();
                        ProductionItem.Status = sqlReader["Status"].ToString();
                        ProductionItem.Level = sqlReader["Level"].ToString();
                        ProductionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        ProductionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        ProductionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        ProductionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        ProductionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        ProductionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        ProductionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        ProductionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        ProductionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        ProductionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        ProductionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        ProductionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        ProductionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        ProductionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        ProductionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        ProductionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        ProductionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        ProductionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        ProductionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        ProductionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        ProductionItem.Remarks = sqlReader["Remarks"].ToString();
                        ProductionItem.Query = sqlReader["Query"].ToString();
                        ProductionItem.IsMovedToQC = sqlReader["IsMovedToQC"].ToString();
                        ProductionItem.TotalRowsCount = Convert.ToInt32(sqlReader["TotalCount"]);       //no need of this in excel export code

                        ProductionItemList.Add(ProductionItem);
                        #endregion
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProductionItemList);

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

        #region Read UNSPSC Versions from Project Noun Modifier
        [HttpGet]
        [Route("ReadUNSPSCVersionsFromProjectNounModifier")]
        public IHttpActionResult ReadUNSPSCVersionsFromProjectNounModifier(string CustomerCode, string ProjectCode, string Noun, string Modifier)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold UNSPSC Versions from Project MRO Dictionary
                List<String> UNSPSCVersionsList = new List<String>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCVersionsOfProjectNounModifier";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    #endregion

                    //Call sp to get the list of UNSPSC Versions
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                        UNSPSCVersionsList.Add(sqlReader["UNSPSCVersionCodeCategory"].ToString());
                    conn.Close();

                    //return list to the request
                    return Ok(UNSPSCVersionsList);
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

        #region Read Production User's Rejected Items of Project  by QC
        [HttpGet]
        [Route("ReadProductionUsersRejectedItemsOfProject")]
        public IHttpActionResult ReadProductionUsersRejectedItemsOfProject(string ProductionUser, string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Production Item List
                List<ProductionItem> productionItemList = new List<ProductionItem>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersRejectedItems";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

                    //Call sp to get Production rejected items of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem productionItem = new ProductionItem();

                        productionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        productionItem.QCItemID = Convert.ToInt64(sqlReader["QCItemID"]);
                        productionItem.QCTestNo = Convert.ToString(sqlReader["QCTestNo"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.ProductionUser = sqlReader["ProductionUser"].ToString();
                        productionItem.QCUser = sqlReader["QCUser"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.QCStatus = sqlReader["QCStatus"].ToString();
                        productionItem.IsMovedToQC = sqlReader["IsMovedToQC"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Query = sqlReader["Query"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();

                        productionItemList.Add(productionItem);
                        #endregion
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(productionItemList);
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

        #region Export the List of Production Rejected Items to Excel
        [HttpGet]
        [Route("ExportProductionRejectedItemsListToExcel")]
        public HttpResponseMessage ExportProductionRejectedItemsListToExcel(string ProductionUser, string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                System.Data.Common.DbDataReader sqlReader;
                string FileName = string.Empty;

                if (!string.IsNullOrEmpty(BatchNo))
                    FileName = "'" + ProductionUser + "' User's Production Rejected Items List from " + CustomerCode + " - " + ProjectCode + " - " + BatchNo + " Project.xlsx";
                else
                    FileName = "'" + ProductionUser + "' User's Production Rejected Items List from " + CustomerCode + " - " + ProjectCode + " Project.xlsx";

                DataTable dtDetails = new DataTable();
                List<ProductionItem> productionItemList = new List<ProductionItem>();
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionUsersRejectedItems";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

                    //Call sp to get Production rejected items of User
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Add Item Details To List
                        ProductionItem productionItem = new ProductionItem();

                        productionItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        productionItem.QCTestNo = Convert.ToString(sqlReader["ProductionItemID"]);
                        productionItem.UniqueID = sqlReader["UniqueID"].ToString();
                        productionItem.ShortDescription = sqlReader["ShortDescription"].ToString();
                        productionItem.LongDescription = sqlReader["LongDescription"].ToString();
                        productionItem.UOM = sqlReader["UOM"].ToString();
                        productionItem.ProductionUser = sqlReader["ProductionUser"].ToString();
                        productionItem.QCUser = sqlReader["QCUser"].ToString();
                        productionItem.MFRName = sqlReader["MFRName"].ToString();
                        productionItem.MFRPN = sqlReader["MFRPN"].ToString();
                        productionItem.VendorName = sqlReader["VendorName"].ToString();
                        productionItem.VendorPN = sqlReader["VendorPN"].ToString();
                        productionItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        productionItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        productionItem.MissingWords = sqlReader["MissingWords"].ToString();
                        productionItem.Noun = sqlReader["Noun"].ToString();
                        productionItem.Modifier = sqlReader["Modifier"].ToString();
                        productionItem.Status = sqlReader["Status"].ToString();
                        productionItem.QCStatus = sqlReader["QCStatus"].ToString();
                        productionItem.IsMovedToQC = sqlReader["IsMovedToQC"].ToString();
                        productionItem.Level = sqlReader["Level"].ToString();
                        productionItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        productionItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        productionItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        productionItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        productionItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        productionItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        productionItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        productionItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        productionItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        productionItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        productionItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        productionItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        productionItem.AdditionalInfo = sqlReader["AdditionalInfo"].ToString();
                        productionItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        productionItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        productionItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        productionItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        productionItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        productionItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        productionItem.PDFURL = sqlReader["PDFURL"].ToString();
                        productionItem.Remarks = sqlReader["Remarks"].ToString();
                        productionItem.Query = sqlReader["Query"].ToString();
                        productionItem.Application = sqlReader["Application"].ToString();
                        productionItem.DWG = sqlReader["DWG"].ToString();
                        productionItem.POS = sqlReader["POS"].ToString();
                        productionItem.ItemNo = sqlReader["ItemNo"].ToString();
                        productionItem.SerialNo = sqlReader["SerialNo"].ToString();
                        productionItem.OtherNo = sqlReader["OtherNo"].ToString();
                        productionItem.KKSCode = sqlReader["KKSCode"].ToString();
                        productionItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        productionItem.BOM = sqlReader["BOM"].ToString();
                        productionItem.GreenItems = sqlReader["GreenItems"].ToString();

                        productionItemList.Add(productionItem);
                        #endregion
                    }
                    conn.Close();

                    if (productionItemList.Count > 0)
                    {
                        if (HttpContext.Current == null)
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);

                        //Create HTTP Response.
                        HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                        #region Setting up the workbook
                        Workbook wb = new Workbook();
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
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

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("UniqueID");
                        ws.Cells[0, 2].PutValue("Short Description");
                        ws.Cells[0, 3].PutValue("Long Description");
                        ws.Cells[0, 4].PutValue("UOM");
                        ws.Cells[0, 5].PutValue("Production User");
                        ws.Cells[0, 6].PutValue("QC User");
                        ws.Cells[0, 7].PutValue("QC Status");
                        ws.Cells[0, 8].PutValue("QC Test No.");
                        ws.Cells[0, 9].PutValue("MFR Name");
                        ws.Cells[0, 10].PutValue("MFR P/N");
                        ws.Cells[0, 11].PutValue("Vendor Name");
                        ws.Cells[0, 12].PutValue("Vendor P/N");
                        ws.Cells[0, 13].PutValue("New Short Description");
                        ws.Cells[0, 14].PutValue("New Long Description");
                        ws.Cells[0, 15].PutValue("Missing Words");
                        ws.Cells[0, 16].PutValue("Noun");
                        ws.Cells[0, 17].PutValue("Modifier");
                        ws.Cells[0, 18].PutValue("Level");
                        ws.Cells[0, 19].PutValue("MFR Name1");
                        ws.Cells[0, 20].PutValue("MFR PN1");
                        ws.Cells[0, 21].PutValue("MFR Name2");
                        ws.Cells[0, 22].PutValue("MFR PN2");
                        ws.Cells[0, 23].PutValue("MFR Name3");
                        ws.Cells[0, 24].PutValue("MFR PN3");
                        ws.Cells[0, 25].PutValue("Vendor Name1");
                        ws.Cells[0, 26].PutValue("Vendor PN1");
                        ws.Cells[0, 27].PutValue("Vendor Name2");
                        ws.Cells[0, 28].PutValue("Vendor PN2");
                        ws.Cells[0, 29].PutValue("Vendor Name3");
                        ws.Cells[0, 30].PutValue("Vendor PN3");
                        ws.Cells[0, 31].PutValue("Additional Info");
                        ws.Cells[0, 32].PutValue("Additional Info From Web");
                        ws.Cells[0, 33].PutValue("UNSPSC Code");
                        ws.Cells[0, 34].PutValue("UNSPSC Category");
                        ws.Cells[0, 35].PutValue("Web Ref URL1");
                        ws.Cells[0, 36].PutValue("Web Ref URL2");
                        ws.Cells[0, 37].PutValue("Web Ref URL3");
                        ws.Cells[0, 38].PutValue("PDF URL");
                        ws.Cells[0, 39].PutValue("Remarks");
                        ws.Cells[0, 40].PutValue("Query");
                        ws.Cells[0, 41].PutValue("Application");
                        ws.Cells[0, 42].PutValue("DWG");
                        ws.Cells[0, 43].PutValue("POS");
                        ws.Cells[0, 44].PutValue("ItemNo");
                        ws.Cells[0, 45].PutValue("SerialNo");
                        ws.Cells[0, 46].PutValue("OtherNo");
                        ws.Cells[0, 47].PutValue("KKSCode");
                        ws.Cells[0, 48].PutValue("Assembly / Part");
                        ws.Cells[0, 49].PutValue("BOM");
                        ws.Cells[0, 50].PutValue("GreenItems");

                        for (int c = 0; c <= 50; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (ProductionItem pi in productionItemList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(pi.UniqueID);
                            ws.Cells[row, 2].PutValue(pi.ShortDescription);
                            ws.Cells[row, 3].PutValue(pi.LongDescription);
                            ws.Cells[row, 4].PutValue(pi.UOM);
                            ws.Cells[row, 5].PutValue(pi.ProductionUser);
                            ws.Cells[row, 6].PutValue(pi.QCUser);
                            ws.Cells[row, 7].PutValue(pi.QCStatus);
                            ws.Cells[row, 8].PutValue(pi.QCTestNo);
                            ws.Cells[row, 9].PutValue(pi.MFRName);
                            ws.Cells[row, 10].PutValue(pi.MFRPN);
                            ws.Cells[row, 11].PutValue(pi.VendorName);
                            ws.Cells[row, 12].PutValue(pi.VendorPN);
                            ws.Cells[row, 13].PutValue(pi.NewShortDescription);
                            ws.Cells[row, 14].PutValue(pi.NewLongDescription);
                            ws.Cells[row, 15].PutValue(pi.MissingWords);
                            ws.Cells[row, 16].PutValue(pi.Noun);
                            ws.Cells[row, 17].PutValue(pi.Modifier);
                            ws.Cells[row, 18].PutValue(pi.Level);
                            ws.Cells[row, 19].PutValue(pi.MFRName1);
                            ws.Cells[row, 20].PutValue(pi.MFRPN1);
                            ws.Cells[row, 21].PutValue(pi.MFRName2);
                            ws.Cells[row, 22].PutValue(pi.MFRPN2);
                            ws.Cells[row, 23].PutValue(pi.MFRName3);
                            ws.Cells[row, 24].PutValue(pi.MFRPN3);
                            ws.Cells[row, 25].PutValue(pi.VendorName1);
                            ws.Cells[row, 26].PutValue(pi.VendorPN1);
                            ws.Cells[row, 27].PutValue(pi.VendorName2);
                            ws.Cells[row, 28].PutValue(pi.VendorPN2);
                            ws.Cells[row, 29].PutValue(pi.VendorName3);
                            ws.Cells[row, 30].PutValue(pi.VendorPN3);
                            ws.Cells[row, 31].PutValue(pi.AdditionalInfo);
                            ws.Cells[row, 32].PutValue(pi.AdditionalInfoFromWeb);
                            ws.Cells[row, 33].PutValue(pi.UNSPSCCode);
                            ws.Cells[row, 34].PutValue(pi.UNSPSCCategory);
                            ws.Cells[row, 35].PutValue(pi.WebRefURL1);
                            ws.Cells[row, 36].PutValue(pi.WebRefURL2);
                            ws.Cells[row, 37].PutValue(pi.WebRefURL3);
                            ws.Cells[row, 38].PutValue(pi.PDFURL);
                            ws.Cells[row, 39].PutValue(pi.Remarks);
                            ws.Cells[row, 40].PutValue(pi.Query);
                            ws.Cells[row, 41].PutValue(pi.Application);
                            ws.Cells[row, 42].PutValue(pi.DWG);
                            ws.Cells[row, 43].PutValue(pi.POS);
                            ws.Cells[row, 44].PutValue(pi.ItemNo);
                            ws.Cells[row, 45].PutValue(pi.SerialNo);
                            ws.Cells[row, 46].PutValue(pi.OtherNo);
                            ws.Cells[row, 47].PutValue(pi.KKSCode);
                            ws.Cells[row, 48].PutValue(pi.AssemblyOrPart);
                            ws.Cells[row, 49].PutValue(pi.BOM);
                            ws.Cells[row, 50].PutValue(pi.GreenItems);
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
                            ws.Cells[row, 9].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 11].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 12].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 14].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 15].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 19].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 21].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 23].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 25].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 27].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 28].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 29].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 31].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 32].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 33].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 35].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 36].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 37].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 38].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 39].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 40].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 41].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 42].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 47].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 48].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 49].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 50].SetStyle(styleCenterAlignData);
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
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Read Project Overall Status
        [HttpGet]
        [Route("ReadProjectOverallStatus")]
        public HttpResponseMessage ReadProjectOverallStatus(string CustomerCode, string ProjectCode, string BatchNo = "", string Status = "P")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                DataSet dataSet = new DataSet();
                dataSet = _BLLProduction.ReadProjectOverallStatusAllResults(CustomerCode, ProjectCode, BatchNo, Status);

                ProjectOverallStatusCount projectOverallStatusCount = new ProjectOverallStatusCount();
                List<ProjectStatusUserDateCount> projectStatusUserDateCountList = new List<ProjectStatusUserDateCount>();
                List<ProjectStatusSKUs> projectStatusSKUsList = new List<ProjectStatusSKUs>();

                DataTable dtStatusCount = new DataTable();
                DataTable dtUserDateCount = new DataTable();
                DataTable dtProjectStatusSKUs = new DataTable();

                if (dataSet.Tables.Count == 3)
                {
                    dtStatusCount = dataSet.Tables[0];

                    if (dtStatusCount.Rows.Count > 0)
                    {
                        projectOverallStatusCount.Volume = Convert.ToInt32(dtStatusCount.Rows[0]["Volume"]);
                        projectOverallStatusCount.AllocatedCount = Convert.ToInt32(dtStatusCount.Rows[0]["ProductionAllocatedCount"]);
                        projectOverallStatusCount.YetToAllocate = Convert.ToInt32(dtStatusCount.Rows[0]["YetToAllocate"]);
                        projectOverallStatusCount.Processed = Convert.ToInt32(dtStatusCount.Rows[0]["Processed"]);
                        projectOverallStatusCount.QCApproved = Convert.ToInt32(dtStatusCount.Rows[0]["QCApproved"]);
                        projectOverallStatusCount.Query = Convert.ToInt32(dtStatusCount.Rows[0]["Query"]);
                        projectOverallStatusCount.ProductionCompletedPercentage = Convert.ToDecimal(dtStatusCount.Rows[0]["ProductionCompletedPercentage"]);
                        projectOverallStatusCount.QCCompletedPercentage = Convert.ToDecimal(dtStatusCount.Rows[0]["QCCompletedPercentage"]);
                    }

                    dtStatusCount = dataSet.Tables[1];

                    foreach (DataRow dr in dtStatusCount.Rows)
                    {
                        ProjectStatusUserDateCount projectStatusUserDateCount = new ProjectStatusUserDateCount();
                        projectStatusUserDateCount.User = dr["User"].ToString();
                        projectStatusUserDateCount.UpdatedOn = Convert.ToDateTime(dr["UpdatedOn"]);
                        projectStatusUserDateCount.Count = Convert.ToInt32(dr["CountOfItems"]);
                        projectStatusUserDateCountList.Add(projectStatusUserDateCount);
                    }

                    dtProjectStatusSKUs = dataSet.Tables[2];

                    foreach (DataRow dr in dtProjectStatusSKUs.Rows)
                    {
                        ProjectStatusSKUs projectStatusSKUs = new ProjectStatusSKUs();
                        projectStatusSKUs.UniqueID = dr["UniqueID"].ToString();
                        projectStatusSKUs.Level = dr["Level"].ToString();
                        projectStatusSKUs.User = dr["User"].ToString();
                        projectStatusSKUs.UpdatedOn = Convert.ToDateTime(dr["UpdatedOn"]);
                        projectStatusSKUsList.Add(projectStatusSKUs);
                    }

                    var result = new TripleResultsDto
                    {
                        projectOverallStatusCount = projectOverallStatusCount,
                        projectStatusUserDateCountList = projectStatusUserDateCountList,
                        projectStatusSKUsList = projectStatusSKUsList,
                    };

                    return Request.CreateResponse(HttpStatusCode.OK, result);

                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Read Date Range Status
        [HttpGet]
        [Route("ReadDateRangeStatus")]
        public HttpResponseMessage ReadDateRangeStatus(DateTime FromDate, DateTime ToDate, string Status = "A")
        {
            try
            {
                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadDateRangeStatus(FromDate, ToDate, Status);

                List<ProductionItem> productionItemsList = new List<ProductionItem>();
                if (dataTable.Rows.Count > 0)
                {
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        ProductionItem productionItem = new ProductionItem();
                        productionItem.UpdatedOn = Convert.ToDateTime(dr["UpdatedOn"]);
                        productionItem.ProductionUser = dr["ProductionUser"].ToString();
                        productionItem.QCUser = dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "";
                        productionItem.CustomerCode = dr["CustomerCode"].ToString();
                        productionItem.ProjectCode = dr["ProjectCode"].ToString();
                        productionItem.BatchNo = dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "";
                        productionItem.UniqueID = dr["UniqueID"].ToString().Trim();
                        productionItem.ShortDescription = dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "";
                        productionItem.LongDescription = dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "";
                        productionItem.UOM = dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "";
                        productionItem.NewShortDescription = dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "";
                        productionItem.NewLongDescription = dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "";
                        productionItem.MissingWords = dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "";
                        productionItem.MFRName = dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "";
                        productionItem.MFRPN = dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "";
                        productionItem.VendorName = dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "";
                        productionItem.VendorPN = dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "";
                        productionItem.Noun = dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "";
                        productionItem.Modifier = dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "";
                        productionItem.Level = dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "";
                        productionItem.MFRName1 = dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "";
                        productionItem.MFRPN1 = dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "";
                        productionItem.MFRName2 = dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "";
                        productionItem.MFRPN2 = dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "";
                        productionItem.MFRName3 = dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "";
                        productionItem.MFRPN3 = dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "";
                        productionItem.VendorName1 = dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "";
                        productionItem.VendorPN1 = dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "";
                        productionItem.VendorName2 = dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "";
                        productionItem.VendorPN2 = dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "";
                        productionItem.VendorName3 = dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "";
                        productionItem.VendorPN3 = dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "";
                        productionItem.AdditionalInfo = dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "";
                        productionItem.AdditionalInfoFromWeb = dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "";
                        productionItem.UNSPSCCode = dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "";
                        productionItem.UNSPSCCategory = dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "";
                        productionItem.WebRefURL1 = dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "";
                        productionItem.WebRefURL2 = dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "";
                        productionItem.WebRefURL3 = dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "";
                        productionItem.PDFURL = dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "";
                        productionItem.Remarks = dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "";
                        productionItem.Query = dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "";
                        productionItem.Application = dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "";
                        productionItem.DWG = dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "";
                        productionItem.POS = dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "";
                        productionItem.ItemNo = dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "";
                        productionItem.SerialNo = dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "";
                        productionItem.OtherNo = dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "";
                        productionItem.KKSCode = dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "";
                        productionItem.AssemblyOrPart = dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "";
                        productionItem.BOM = dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "";
                        productionItem.GreenItems = dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "";
                        productionItem.ProcessedCount = Convert.ToInt32(dr["ProcessedCount"]);
                        productionItem.QCApprovedCount = Convert.ToInt32(dr["QCApprovedCount"]);
                        productionItem.RejectedCount = Convert.ToInt32(dr["RejectedCount"]);
                        productionItem.QueryCount = Convert.ToInt32(dr["QueryCount"]);
                        productionItemsList.Add(productionItem);
                    }

                    //return Request.CreateResponse(HttpStatusCode.OK, productionItemsList);


                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", productionItemsList.Count),
                    new JProperty("ProductionItemsList", new JArray(from i in productionItemsList
                                                                    select new JObject(
                                                                new JProperty("UpdatedOn", i.UpdatedOn),
                                                                new JProperty("ProductionUser", i.ProductionUser),
                                                                new JProperty("QCUser", i.QCUser),
                                                                new JProperty("CustomerCode", i.CustomerCode),
                                                                new JProperty("ProjectCode", i.ProjectCode),
                                                                new JProperty("BatchNo", i.BatchNo),
                                                                new JProperty("UniqueID", i.UniqueID),
                                                                new JProperty("ShortDescription", i.ShortDescription),
                                                                new JProperty("LongDescription", i.LongDescription),
                                                                new JProperty("UOM", i.UOM),
                                                                new JProperty("NewShortDescription", i.NewShortDescription),
                                                                new JProperty("NewLongDescription", i.NewLongDescription),
                                                                new JProperty("MissingWords", i.MissingWords),
                                                                new JProperty("MFRName", i.MFRName),
                                                                new JProperty("MFRPN", i.MFRPN),
                                                                new JProperty("VendorName", i.VendorName),
                                                                new JProperty("VendorPN", i.VendorPN),
                                                                new JProperty("Noun", i.Noun),
                                                                new JProperty("Modifier", i.Modifier),
                                                                new JProperty("Level", i.Level),
                                                                new JProperty("MFRName1", i.MFRName1),
                                                                new JProperty("MFRPN1", i.MFRPN1),
                                                                new JProperty("MFRName2", i.MFRName2),
                                                                new JProperty("MFRPN2", i.MFRPN2),
                                                                new JProperty("MFRName3", i.MFRName3),
                                                                new JProperty("MFRPN3", i.MFRPN3),
                                                                new JProperty("VendorName1", i.VendorName1),
                                                                new JProperty("VendorPN1", i.VendorPN1),
                                                                new JProperty("VendorName2", i.VendorName2),
                                                                new JProperty("VendorPN2", i.VendorPN2),
                                                                new JProperty("VendorName3", i.VendorName3),
                                                                new JProperty("VendorPN3", i.VendorPN3),
                                                                new JProperty("AdditionalInfoFromInput", i.AdditionalInfo),
                                                                new JProperty("AdditionalInfoFromWeb", i.AdditionalInfoFromWeb),
                                                                new JProperty("UNSPSCCode", i.UNSPSCCode),
                                                                new JProperty("UNSPSCCategory", i.UNSPSCCategory),
                                                                new JProperty("WebRefURL1", i.WebRefURL1),
                                                                new JProperty("WebRefURL2", i.WebRefURL2),
                                                                new JProperty("WebRefURL3", i.WebRefURL3),
                                                                new JProperty("PDFURL", i.PDFURL),
                                                                new JProperty("Remarks", i.Remarks),
                                                                new JProperty("Query", i.Query),
                                                                new JProperty("Application", i.Application),
                                                                new JProperty("DWG", i.DWG),
                                                                new JProperty("POS", i.POS),
                                                                new JProperty("ItemNo", i.ItemNo),
                                                                new JProperty("SerialNo", i.SerialNo),
                                                                new JProperty("OtherNo", i.OtherNo),
                                                                new JProperty("KKSCode", i.KKSCode),
                                                                new JProperty("AssemblyOrPart", i.AssemblyOrPart),
                                                                new JProperty("BOM", i.BOM),
                                                                new JProperty("GreenItems", i.GreenItems)
                                                                ))),
                    new JProperty("ProcessedCount", productionItemsList[0].ProcessedCount),
                    new JProperty("QCApprovedCount", productionItemsList[0].QCApprovedCount),
                    new JProperty("RejectedCount", productionItemsList[0].RejectedCount),
                    new JProperty("QueryCount", productionItemsList[0].QueryCount)
                    );
                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }

            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export Date Range Status To Excel
        [HttpGet]
        [Route("ExportDateRangeStatusListToExcel")]
        public HttpResponseMessage ExportDateRangeStatusListToExcel(DateTime FromDate, DateTime ToDate, string UserID, string Status = "A")
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Date Range Status Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    DataTable dataTable = new DataTable();
                    dataTable = _BLLProduction.ReadDateRangeStatus(FromDate, ToDate, Status);

                    if (dataTable.Rows.Count > 0)
                    {
                        if (HttpContext.Current == null)
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);

                        //Create HTTP Response.
                        HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                        string FileName = "Date Range Status.xlsx";

                        #region Setting up the workbook
                        AsposeHelpers asposeHelpers = new AsposeHelpers();
                        Workbook wb = asposeHelpers.GetWorkbook();
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                        var ws = wb.Worksheets[0];
                        int row = 1;
                        #endregion

                        #region Setting Styles
                        Aspose.Cells.Style styleHeader = asposeHelpers.GetStyle(wb, 0, "header");
                        Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");
                        Aspose.Cells.Style styleLeftAlignData = asposeHelpers.GetStyle(wb, 0, "left");
                        #endregion

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Date");
                        ws.Cells[0, 2].PutValue("Production User");
                        ws.Cells[0, 3].PutValue("QC User");
                        ws.Cells[0, 4].PutValue("Customer Code");
                        ws.Cells[0, 5].PutValue("Project Code");
                        ws.Cells[0, 6].PutValue("Batch No.");
                        ws.Cells[0, 7].PutValue("UniqueID");
                        ws.Cells[0, 8].PutValue("Short Description");
                        ws.Cells[0, 9].PutValue("Long Description");
                        ws.Cells[0, 10].PutValue("UOM");
                        ws.Cells[0, 11].PutValue("New Short Description");
                        ws.Cells[0, 12].PutValue("New Long Description");
                        ws.Cells[0, 13].PutValue("Missing Words");
                        ws.Cells[0, 14].PutValue("MFR Name");
                        ws.Cells[0, 15].PutValue("MFR P/N");
                        ws.Cells[0, 16].PutValue("Vendor Name");
                        ws.Cells[0, 17].PutValue("Vendor P/N");
                        ws.Cells[0, 18].PutValue("Noun");
                        ws.Cells[0, 19].PutValue("Modifier");
                        ws.Cells[0, 20].PutValue("Level");
                        ws.Cells[0, 21].PutValue("MFR Name1");
                        ws.Cells[0, 22].PutValue("MFR PN1");
                        ws.Cells[0, 23].PutValue("MFR Name2");
                        ws.Cells[0, 24].PutValue("MFR PN2");
                        ws.Cells[0, 25].PutValue("MFR Name3");
                        ws.Cells[0, 26].PutValue("MFR PN3");
                        ws.Cells[0, 27].PutValue("Vendor Name1");
                        ws.Cells[0, 28].PutValue("Vendor PN1");
                        ws.Cells[0, 29].PutValue("Vendor Name2");
                        ws.Cells[0, 30].PutValue("Vendor PN2");
                        ws.Cells[0, 31].PutValue("Vendor Name3");
                        ws.Cells[0, 32].PutValue("Vendor PN3");
                        ws.Cells[0, 33].PutValue("Additional Info");
                        ws.Cells[0, 34].PutValue("Additional Info From Web");
                        ws.Cells[0, 35].PutValue("UNSPSC Code");
                        ws.Cells[0, 36].PutValue("UNSPSC Category");
                        ws.Cells[0, 37].PutValue("Web Ref URL1");
                        ws.Cells[0, 38].PutValue("Web Ref URL2");
                        ws.Cells[0, 39].PutValue("Web Ref URL3");
                        ws.Cells[0, 40].PutValue("PDF URL");
                        ws.Cells[0, 41].PutValue("Remarks");
                        ws.Cells[0, 42].PutValue("Query");
                        ws.Cells[0, 43].PutValue("Application");
                        ws.Cells[0, 44].PutValue("DWG");
                        ws.Cells[0, 45].PutValue("POS");
                        ws.Cells[0, 46].PutValue("ItemNo");
                        ws.Cells[0, 47].PutValue("SerialNo");
                        ws.Cells[0, 48].PutValue("OtherNo");
                        ws.Cells[0, 49].PutValue("KKSCode");
                        ws.Cells[0, 50].PutValue("Assembly / Part");
                        ws.Cells[0, 51].PutValue("BOM");
                        ws.Cells[0, 52].PutValue("GreenItems");

                        for (int c = 0; c <= 52; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(Convert.ToDateTime(dr["UpdatedOn"]).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 2].PutValue(dr["ProductionUser"].ToString());
                            ws.Cells[row, 3].PutValue(dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "");
                            ws.Cells[row, 4].PutValue(dr["CustomerCode"].ToString());
                            ws.Cells[row, 5].PutValue(dr["ProjectCode"].ToString());
                            ws.Cells[row, 6].PutValue(dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "");
                            ws.Cells[row, 7].PutValue(dr["UniqueID"].ToString().Trim());
                            ws.Cells[row, 8].PutValue(dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 9].PutValue(dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 10].PutValue(dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "");
                            ws.Cells[row, 11].PutValue(dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 12].PutValue(dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 13].PutValue(dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "");
                            ws.Cells[row, 14].PutValue(dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "");
                            ws.Cells[row, 15].PutValue(dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "");
                            ws.Cells[row, 16].PutValue(dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "");
                            ws.Cells[row, 17].PutValue(dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "");
                            ws.Cells[row, 18].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                            ws.Cells[row, 19].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                            ws.Cells[row, 20].PutValue(dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "");
                            ws.Cells[row, 21].PutValue(dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "");
                            ws.Cells[row, 22].PutValue(dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "");
                            ws.Cells[row, 23].PutValue(dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "");
                            ws.Cells[row, 24].PutValue(dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "");
                            ws.Cells[row, 25].PutValue(dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "");
                            ws.Cells[row, 26].PutValue(dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "");
                            ws.Cells[row, 27].PutValue(dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "");
                            ws.Cells[row, 28].PutValue(dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "");
                            ws.Cells[row, 29].PutValue(dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "");
                            ws.Cells[row, 30].PutValue(dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "");
                            ws.Cells[row, 31].PutValue(dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "");
                            ws.Cells[row, 32].PutValue(dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "");
                            ws.Cells[row, 33].PutValue(dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "");
                            ws.Cells[row, 34].PutValue(dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "");
                            ws.Cells[row, 35].PutValue(dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "");
                            ws.Cells[row, 36].PutValue(dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "");
                            ws.Cells[row, 37].PutValue(dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "");
                            ws.Cells[row, 38].PutValue(dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "");
                            ws.Cells[row, 39].PutValue(dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "");
                            ws.Cells[row, 40].PutValue(dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "");
                            ws.Cells[row, 41].PutValue(dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "");
                            ws.Cells[row, 42].PutValue(dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "");
                            ws.Cells[row, 43].PutValue(dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "");
                            ws.Cells[row, 44].PutValue(dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "");
                            ws.Cells[row, 45].PutValue(dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "");
                            ws.Cells[row, 46].PutValue(dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "");
                            ws.Cells[row, 47].PutValue(dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "");
                            ws.Cells[row, 48].PutValue(dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "");
                            ws.Cells[row, 49].PutValue(dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "");
                            ws.Cells[row, 50].PutValue(dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "");
                            ws.Cells[row, 51].PutValue(dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "");
                            ws.Cells[row, 52].PutValue(dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "");
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 8].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 9].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 11].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 12].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 25].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 27].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 28].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 29].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 31].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 32].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 33].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 35].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 36].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 37].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 38].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 39].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 40].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 41].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 42].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 47].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 48].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 49].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 50].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 51].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 52].SetStyle(styleCenterAlignData);
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
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region User Based Status Report
        #region Read Production and QC unique User Names
        [HttpGet]
        [Route("ReadProductionAndQCUniqueUserNames")]
        public HttpResponseMessage ReadProductionAndQCUniqueUserNames()
        {
            try
            {
                List<string> userNames = new List<string>();

                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadProductionAndQCUserNames();

                if (dataTable.Rows.Count > 0)
                {
                    foreach (DataRow dr in dataTable.Rows)
                        userNames.Add(dr["UserName"].ToString());

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                                                    new JProperty("UserNamesList", new JArray(from i in userNames
                                                                                              select new JObject(
                                                                                          new JProperty("UserName", i)))));
                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }

        #endregion

        #region Read User Based Status
        [HttpGet]
        [Route("ReadUserBasedStatus")]
        public HttpResponseMessage ReadUserBasedStatus(string UserName, string Status = "A")
        {
            try
            {
                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadUserBasedStatus(UserName, Status);

                List<ProductionItem> productionItemsList = new List<ProductionItem>();
                if (dataTable.Rows.Count > 0)
                {
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        ProductionItem productionItem = new ProductionItem();
                        productionItem.UpdatedOn = Convert.ToDateTime(dr["UpdatedOn"]);
                        productionItem.ProductionUser = dr["ProductionUser"].ToString();
                        productionItem.QCUser = dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "";
                        productionItem.CustomerCode = dr["CustomerCode"].ToString();
                        productionItem.ProjectCode = dr["ProjectCode"].ToString();
                        productionItem.BatchNo = dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "";
                        productionItem.UniqueID = dr["UniqueID"].ToString().Trim();
                        productionItem.ShortDescription = dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "";
                        productionItem.LongDescription = dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "";
                        productionItem.UOM = dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "";
                        productionItem.NewShortDescription = dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "";
                        productionItem.NewLongDescription = dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "";
                        productionItem.MissingWords = dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "";
                        productionItem.MFRName = dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "";
                        productionItem.MFRPN = dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "";
                        productionItem.VendorName = dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "";
                        productionItem.VendorPN = dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "";
                        productionItem.Noun = dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "";
                        productionItem.Modifier = dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "";
                        productionItem.Level = dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "";
                        productionItem.MFRName1 = dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "";
                        productionItem.MFRPN1 = dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "";
                        productionItem.MFRName2 = dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "";
                        productionItem.MFRPN2 = dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "";
                        productionItem.MFRName3 = dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "";
                        productionItem.MFRPN3 = dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "";
                        productionItem.VendorName1 = dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "";
                        productionItem.VendorPN1 = dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "";
                        productionItem.VendorName2 = dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "";
                        productionItem.VendorPN2 = dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "";
                        productionItem.VendorName3 = dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "";
                        productionItem.VendorPN3 = dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "";
                        productionItem.AdditionalInfo = dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "";
                        productionItem.AdditionalInfoFromWeb = dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "";
                        productionItem.UNSPSCCode = dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "";
                        productionItem.UNSPSCCategory = dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "";
                        productionItem.WebRefURL1 = dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "";
                        productionItem.WebRefURL2 = dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "";
                        productionItem.WebRefURL3 = dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "";
                        productionItem.PDFURL = dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "";
                        productionItem.Remarks = dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "";
                        productionItem.Query = dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "";
                        productionItem.Application = dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "";
                        productionItem.DWG = dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "";
                        productionItem.POS = dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "";
                        productionItem.ItemNo = dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "";
                        productionItem.SerialNo = dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "";
                        productionItem.OtherNo = dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "";
                        productionItem.KKSCode = dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "";
                        productionItem.AssemblyOrPart = dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "";
                        productionItem.BOM = dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "";
                        productionItem.GreenItems = dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "";
                        productionItem.AllocatedCount = Convert.ToInt32(dr["AllocatedCount"]);
                        productionItem.ProcessedCount = Convert.ToInt32(dr["ProcessedCount"]);
                        productionItem.PendingCount = Convert.ToInt32(dr["PendingCount"]);
                        productionItem.QCApprovedCount = Convert.ToInt32(dr["QCApprovedCount"]);
                        productionItem.RejectedCount = Convert.ToInt32(dr["RejectedCount"]);
                        productionItem.QueryCount = Convert.ToInt32(dr["QueryCount"]);
                        productionItemsList.Add(productionItem);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", productionItemsList.Count),
                    new JProperty("ProductionItemsList", new JArray(from i in productionItemsList
                                                                    select new JObject(
                                                                new JProperty("UpdatedOn", i.UpdatedOn),
                                                                new JProperty("ProductionUser", i.ProductionUser),
                                                                new JProperty("QCUser", i.QCUser),
                                                                new JProperty("CustomerCode", i.CustomerCode),
                                                                new JProperty("ProjectCode", i.ProjectCode),
                                                                new JProperty("BatchNo", i.BatchNo),
                                                                new JProperty("UniqueID", i.UniqueID),
                                                                new JProperty("ShortDescription", i.ShortDescription),
                                                                new JProperty("LongDescription", i.LongDescription),
                                                                new JProperty("UOM", i.UOM),
                                                                new JProperty("NewShortDescription", i.NewShortDescription),
                                                                new JProperty("NewLongDescription", i.NewLongDescription),
                                                                new JProperty("MissingWords", i.MissingWords),
                                                                new JProperty("MFRName", i.MFRName),
                                                                new JProperty("MFRPN", i.MFRPN),
                                                                new JProperty("VendorName", i.VendorName),
                                                                new JProperty("VendorPN", i.VendorPN),
                                                                new JProperty("Noun", i.Noun),
                                                                new JProperty("Modifier", i.Modifier),
                                                                new JProperty("Level", i.Level),
                                                                new JProperty("MFRName1", i.MFRName1),
                                                                new JProperty("MFRPN1", i.MFRPN1),
                                                                new JProperty("MFRName2", i.MFRName2),
                                                                new JProperty("MFRPN2", i.MFRPN2),
                                                                new JProperty("MFRName3", i.MFRName3),
                                                                new JProperty("MFRPN3", i.MFRPN3),
                                                                new JProperty("VendorName1", i.VendorName1),
                                                                new JProperty("VendorPN1", i.VendorPN1),
                                                                new JProperty("VendorName2", i.VendorName2),
                                                                new JProperty("VendorPN2", i.VendorPN2),
                                                                new JProperty("VendorName3", i.VendorName3),
                                                                new JProperty("VendorPN3", i.VendorPN3),
                                                                new JProperty("AdditionalInfoFromInput", i.AdditionalInfo),
                                                                new JProperty("AdditionalInfoFromWeb", i.AdditionalInfoFromWeb),
                                                                new JProperty("UNSPSCCode", i.UNSPSCCode),
                                                                new JProperty("UNSPSCCategory", i.UNSPSCCategory),
                                                                new JProperty("WebRefURL1", i.WebRefURL1),
                                                                new JProperty("WebRefURL2", i.WebRefURL2),
                                                                new JProperty("WebRefURL3", i.WebRefURL3),
                                                                new JProperty("PDFURL", i.PDFURL),
                                                                new JProperty("Remarks", i.Remarks),
                                                                new JProperty("Query", i.Query),
                                                                new JProperty("Application", i.Application),
                                                                new JProperty("DWG", i.DWG),
                                                                new JProperty("POS", i.POS),
                                                                new JProperty("ItemNo", i.ItemNo),
                                                                new JProperty("SerialNo", i.SerialNo),
                                                                new JProperty("OtherNo", i.OtherNo),
                                                                new JProperty("KKSCode", i.KKSCode),
                                                                new JProperty("AssemblyOrPart", i.AssemblyOrPart),
                                                                new JProperty("BOM", i.BOM),
                                                                new JProperty("GreenItems", i.GreenItems)
                                                                ))),
                    new JProperty("AllocatedCount", productionItemsList[0].AllocatedCount),
                    new JProperty("ProcessedCount", productionItemsList[0].ProcessedCount),
                    new JProperty("PendingCount", productionItemsList[0].PendingCount),
                    new JProperty("QCApprovedCount", productionItemsList[0].QCApprovedCount),
                    new JProperty("RejectedCount", productionItemsList[0].RejectedCount),
                    new JProperty("QueryCount", productionItemsList[0].QueryCount)
                    );

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export User Based Status To Excel
        [HttpGet]
        [Route("ExportUserBasedStatusListToExcel")]
        public HttpResponseMessage ExportUserBasedStatusListToExcel(string UserName, string UserID, string Status = "A")
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("User Based Status Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string StatusLabel = string.Empty;

                    switch (Status)
                    {
                        case "A":
                            StatusLabel = "Allocated";
                            break;
                        case "P":
                            StatusLabel = "Processed";
                            break;
                        case "I":
                            StatusLabel = "Pending";
                            break;
                        case "Q":
                            StatusLabel = "QC Approvd";
                            break;
                        case "R":
                            StatusLabel = "Rejected";
                            break;
                        case "U":
                            StatusLabel = "Query";
                            break;
                    }



                    DataTable dataTable = new DataTable();
                    dataTable = _BLLProduction.ReadUserBasedStatus(UserName, Status);

                    if (dataTable.Rows.Count > 0)
                    {
                        if (HttpContext.Current == null)
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);

                        //Create HTTP Response.
                        HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                        string FileName = "User Based Status Report UserName - " + UserName + " - " + " and Status - " + StatusLabel + ".xlsx";

                        #region Setting up the workbook
                        AsposeHelpers asposeHelpers = new AsposeHelpers();
                        Workbook wb = asposeHelpers.GetWorkbook();
                        var ws = wb.Worksheets[0];
                        int row = 1;
                        #endregion

                        #region Setting Styles
                        Aspose.Cells.Style styleHeader = asposeHelpers.GetStyle(wb, 0, "header");
                        Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");
                        Aspose.Cells.Style styleLeftAlignData = asposeHelpers.GetStyle(wb, 0, "left");
                        #endregion

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Date");
                        ws.Cells[0, 2].PutValue("Production User");
                        ws.Cells[0, 3].PutValue("QC User");
                        ws.Cells[0, 4].PutValue("Customer Code");
                        ws.Cells[0, 5].PutValue("Project Code");
                        ws.Cells[0, 6].PutValue("Batch No.");
                        ws.Cells[0, 7].PutValue("UniqueID");
                        ws.Cells[0, 8].PutValue("Short Description");
                        ws.Cells[0, 9].PutValue("Long Description");
                        ws.Cells[0, 10].PutValue("UOM");
                        ws.Cells[0, 11].PutValue("New Short Description");
                        ws.Cells[0, 12].PutValue("New Long Description");
                        ws.Cells[0, 13].PutValue("Missing Words");
                        ws.Cells[0, 14].PutValue("MFR Name");
                        ws.Cells[0, 15].PutValue("MFR P/N");
                        ws.Cells[0, 16].PutValue("Vendor Name");
                        ws.Cells[0, 17].PutValue("Vendor P/N");
                        ws.Cells[0, 18].PutValue("Noun");
                        ws.Cells[0, 19].PutValue("Modifier");
                        ws.Cells[0, 20].PutValue("Level");
                        ws.Cells[0, 21].PutValue("MFR Name1");
                        ws.Cells[0, 22].PutValue("MFR PN1");
                        ws.Cells[0, 23].PutValue("MFR Name2");
                        ws.Cells[0, 24].PutValue("MFR PN2");
                        ws.Cells[0, 25].PutValue("MFR Name3");
                        ws.Cells[0, 26].PutValue("MFR PN3");
                        ws.Cells[0, 27].PutValue("Vendor Name1");
                        ws.Cells[0, 28].PutValue("Vendor PN1");
                        ws.Cells[0, 29].PutValue("Vendor Name2");
                        ws.Cells[0, 30].PutValue("Vendor PN2");
                        ws.Cells[0, 31].PutValue("Vendor Name3");
                        ws.Cells[0, 32].PutValue("Vendor PN3");
                        ws.Cells[0, 33].PutValue("Additional Info");
                        ws.Cells[0, 34].PutValue("Additional Info From Web");
                        ws.Cells[0, 35].PutValue("UNSPSC Code");
                        ws.Cells[0, 36].PutValue("UNSPSC Category");
                        ws.Cells[0, 37].PutValue("Web Ref URL1");
                        ws.Cells[0, 38].PutValue("Web Ref URL2");
                        ws.Cells[0, 39].PutValue("Web Ref URL3");
                        ws.Cells[0, 40].PutValue("PDF URL");
                        ws.Cells[0, 41].PutValue("Remarks");
                        ws.Cells[0, 42].PutValue("Query");
                        ws.Cells[0, 43].PutValue("Application");
                        ws.Cells[0, 44].PutValue("DWG");
                        ws.Cells[0, 45].PutValue("POS");
                        ws.Cells[0, 46].PutValue("ItemNo");
                        ws.Cells[0, 47].PutValue("SerialNo");
                        ws.Cells[0, 48].PutValue("OtherNo");
                        ws.Cells[0, 49].PutValue("KKSCode");
                        ws.Cells[0, 50].PutValue("Assembly / Part");
                        ws.Cells[0, 51].PutValue("BOM");
                        ws.Cells[0, 52].PutValue("GreenItems");

                        for (int c = 0; c <= 52; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(Convert.ToDateTime(dr["UpdatedOn"]).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 2].PutValue(dr["ProductionUser"].ToString());
                            ws.Cells[row, 3].PutValue(dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "");
                            ws.Cells[row, 4].PutValue(dr["CustomerCode"].ToString());
                            ws.Cells[row, 5].PutValue(dr["ProjectCode"].ToString());
                            ws.Cells[row, 6].PutValue(dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "");
                            ws.Cells[row, 7].PutValue(dr["UniqueID"].ToString().Trim());
                            ws.Cells[row, 8].PutValue(dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 9].PutValue(dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 10].PutValue(dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "");
                            ws.Cells[row, 11].PutValue(dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 12].PutValue(dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 13].PutValue(dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "");
                            ws.Cells[row, 14].PutValue(dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "");
                            ws.Cells[row, 15].PutValue(dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "");
                            ws.Cells[row, 16].PutValue(dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "");
                            ws.Cells[row, 17].PutValue(dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "");
                            ws.Cells[row, 18].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                            ws.Cells[row, 19].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                            ws.Cells[row, 20].PutValue(dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "");
                            ws.Cells[row, 21].PutValue(dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "");
                            ws.Cells[row, 22].PutValue(dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "");
                            ws.Cells[row, 23].PutValue(dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "");
                            ws.Cells[row, 24].PutValue(dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "");
                            ws.Cells[row, 25].PutValue(dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "");
                            ws.Cells[row, 26].PutValue(dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "");
                            ws.Cells[row, 27].PutValue(dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "");
                            ws.Cells[row, 28].PutValue(dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "");
                            ws.Cells[row, 29].PutValue(dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "");
                            ws.Cells[row, 30].PutValue(dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "");
                            ws.Cells[row, 31].PutValue(dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "");
                            ws.Cells[row, 32].PutValue(dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "");
                            ws.Cells[row, 33].PutValue(dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "");
                            ws.Cells[row, 34].PutValue(dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "");
                            ws.Cells[row, 35].PutValue(dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "");
                            ws.Cells[row, 36].PutValue(dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "");
                            ws.Cells[row, 37].PutValue(dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "");
                            ws.Cells[row, 38].PutValue(dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "");
                            ws.Cells[row, 39].PutValue(dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "");
                            ws.Cells[row, 40].PutValue(dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "");
                            ws.Cells[row, 41].PutValue(dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "");
                            ws.Cells[row, 42].PutValue(dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "");
                            ws.Cells[row, 43].PutValue(dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "");
                            ws.Cells[row, 44].PutValue(dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "");
                            ws.Cells[row, 45].PutValue(dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "");
                            ws.Cells[row, 46].PutValue(dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "");
                            ws.Cells[row, 47].PutValue(dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "");
                            ws.Cells[row, 48].PutValue(dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "");
                            ws.Cells[row, 49].PutValue(dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "");
                            ws.Cells[row, 50].PutValue(dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "");
                            ws.Cells[row, 51].PutValue(dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "");
                            ws.Cells[row, 52].PutValue(dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "");
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 8].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 9].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 11].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 12].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 25].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 27].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 28].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 29].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 31].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 32].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 33].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 35].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 36].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 37].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 38].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 39].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 40].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 41].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 42].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 47].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 48].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 49].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 50].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 51].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 52].SetStyle(styleCenterAlignData);
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
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion
        #endregion

        #region Project Level Quality Report
        #region Project Level Quality Report Count Stats
        [HttpGet]
        [Route("ReadProjectLevelQualityReportCountStats")]
        public HttpResponseMessage ReadProjectLevelQualityReportCountStats(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadProjectLevelQualityCountStats(CustomerCode, ProjectCode, BatchNo);

                if (dataTable.Rows.Count > 0)
                {
                    JObject PEReport = new JObject(new JProperty("Volume", dataTable.Rows[0]["Volume"]),
                                                    new JProperty("ProcessedCount", dataTable.Rows[0]["ProcessedCount"]),
                                                    new JProperty("AllocatedCount", dataTable.Rows[0]["AllocatedCount"]),
                                                    new JProperty("QCApprovedCount", dataTable.Rows[0]["QCApprovedCount"]),
                                                    new JProperty("AcceptedPercentage", dataTable.Rows[0]["AcceptedPercentage"]),
                                                    new JProperty("RejectedCount", dataTable.Rows[0]["RejectedCount"]),
                                                    new JProperty("NMChangedCount", dataTable.Rows[0]["NMChangedCount"]),
                                                    new JProperty("NMChangedPercentage", dataTable.Rows[0]["NMChangedPercentage"]),
                                                    new JProperty("MFRorSupplierChangedCount", dataTable.Rows[0]["MFRorSupplierChangedCount"]),
                                                    new JProperty("MFRorSupplierChangedPercentage", dataTable.Rows[0]["MFRorSupplierChangedPercentage"]),
                                                    new JProperty("AttributeChangedCount", dataTable.Rows[0]["AttributeChangedCount"]),
                                                    new JProperty("AttributeChangedPercentage", dataTable.Rows[0]["AttributeChangedPercentage"]),
                                                    new JProperty("ProjectScope", dataTable.Rows[0]["ProjectScope"]));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Project Level Quality Report SKU Details
        [HttpGet]
        [Route("ReadProjectLevelQualityReportSKUDetails")]
        public HttpResponseMessage ReadProjectLevelQualityReportSKUDetails(string CustomerCode, string ProjectCode, string BatchNo = "", string Status = "P")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadProjectLevelQualityReportSKUDetails(CustomerCode, ProjectCode, BatchNo, Status);

                List<ProductionItem> productionItemsList = new List<ProductionItem>();
                if (dataTable.Rows.Count > 0)
                {
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        ProductionItem productionItem = new ProductionItem();
                        productionItem.UniqueID = dr["UniqueID"].ToString().Trim();
                        productionItem.ShortDescription = dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "";
                        productionItem.LongDescription = dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "";
                        productionItem.ProductionUser = dr["ProductionUser"].ToString();
                        productionItem.Status = dr["ProductionStatus"].ToString();
                        productionItem.QCUser = dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "";
                        productionItem.QCStatus = dr["QCStatus"] != DBNull.Value ? dr["QCStatus"].ToString().Trim() : "";
                        productionItem.UOM = dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "";
                        productionItem.MFRName = dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "";
                        productionItem.MFRPN = dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "";
                        productionItem.VendorName = dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "";
                        productionItem.VendorPN = dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "";
                        productionItem.NewShortDescription = dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "";
                        productionItem.NewLongDescription = dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "";
                        productionItem.MissingWords = dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "";
                        productionItem.Noun = dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "";
                        productionItem.Modifier = dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "";
                        productionItem.Level = dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "";
                        productionItem.MFRName1 = dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "";
                        productionItem.MFRPN1 = dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "";
                        productionItem.MFRName2 = dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "";
                        productionItem.MFRPN2 = dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "";
                        productionItem.MFRName3 = dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "";
                        productionItem.MFRPN3 = dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "";
                        productionItem.VendorName1 = dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "";
                        productionItem.VendorPN1 = dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "";
                        productionItem.VendorName2 = dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "";
                        productionItem.VendorPN2 = dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "";
                        productionItem.VendorName3 = dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "";
                        productionItem.VendorPN3 = dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "";
                        productionItem.AdditionalInfo = dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "";
                        productionItem.AdditionalInfoFromWeb = dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "";
                        productionItem.UNSPSCCode = dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "";
                        productionItem.UNSPSCCategory = dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "";
                        productionItem.WebRefURL1 = dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "";
                        productionItem.WebRefURL2 = dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "";
                        productionItem.WebRefURL3 = dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "";
                        productionItem.PDFURL = dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "";
                        productionItem.Remarks = dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "";
                        productionItem.Query = dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "";
                        productionItem.Application = dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "";
                        productionItem.DWG = dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "";
                        productionItem.POS = dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "";
                        productionItem.ItemNo = dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "";
                        productionItem.SerialNo = dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "";
                        productionItem.OtherNo = dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "";
                        productionItem.KKSCode = dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "";
                        productionItem.AssemblyOrPart = dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "";
                        productionItem.BOM = dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "";
                        productionItem.GreenItems = dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "";
                        productionItemsList.Add(productionItem);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", productionItemsList.Count),
                    new JProperty("ProductionItemsList", new JArray(from i in productionItemsList
                                                                    select new JObject(
                                                                new JProperty("UniqueID", i.UniqueID),
                                                                new JProperty("ShortDescription", i.ShortDescription),
                                                                new JProperty("LongDescription", i.LongDescription),
                                                                new JProperty("ProductionUser", i.ProductionUser),
                                                                new JProperty("ProductionStatus", i.Status),
                                                                new JProperty("QCUser", i.QCUser),
                                                                new JProperty("QCStatus", i.QCStatus),
                                                                new JProperty("MFRName", i.MFRName),
                                                                new JProperty("MFRPN", i.MFRPN),
                                                                new JProperty("VendorName", i.VendorName),
                                                                new JProperty("VendorPN", i.VendorPN),
                                                                new JProperty("UOM", i.UOM),
                                                                new JProperty("NewShortDescription", i.NewShortDescription),
                                                                new JProperty("NewLongDescription", i.NewLongDescription),
                                                                new JProperty("MissingWords", i.MissingWords),
                                                                new JProperty("Noun", i.Noun),
                                                                new JProperty("Modifier", i.Modifier),
                                                                new JProperty("Level", i.Level),
                                                                new JProperty("MFRName1", i.MFRName1),
                                                                new JProperty("MFRPN1", i.MFRPN1),
                                                                new JProperty("MFRName2", i.MFRName2),
                                                                new JProperty("MFRPN2", i.MFRPN2),
                                                                new JProperty("MFRName3", i.MFRName3),
                                                                new JProperty("MFRPN3", i.MFRPN3),
                                                                new JProperty("VendorName1", i.VendorName1),
                                                                new JProperty("VendorPN1", i.VendorPN1),
                                                                new JProperty("VendorName2", i.VendorName2),
                                                                new JProperty("VendorPN2", i.VendorPN2),
                                                                new JProperty("VendorName3", i.VendorName3),
                                                                new JProperty("VendorPN3", i.VendorPN3),
                                                                new JProperty("AdditionalInfoFromInput", i.AdditionalInfo),
                                                                new JProperty("AdditionalInfoFromWeb", i.AdditionalInfoFromWeb),
                                                                new JProperty("UNSPSCCode", i.UNSPSCCode),
                                                                new JProperty("UNSPSCCategory", i.UNSPSCCategory),
                                                                new JProperty("WebRefURL1", i.WebRefURL1),
                                                                new JProperty("WebRefURL2", i.WebRefURL2),
                                                                new JProperty("WebRefURL3", i.WebRefURL3),
                                                                new JProperty("PDFURL", i.PDFURL),
                                                                new JProperty("Remarks", i.Remarks),
                                                                new JProperty("Query", i.Query),
                                                                new JProperty("Application", i.Application),
                                                                new JProperty("DWG", i.DWG),
                                                                new JProperty("ItemNo", i.ItemNo),
                                                                new JProperty("POS", i.POS),
                                                                new JProperty("SerialNo", i.SerialNo),
                                                                new JProperty("OtherNo", i.OtherNo),
                                                                new JProperty("KKSCode", i.KKSCode),
                                                                new JProperty("AssemblyOrPart", i.AssemblyOrPart),
                                                                new JProperty("BOM", i.BOM),
                                                                new JProperty("GreenItems", i.GreenItems)
                                                                ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export Project Level Quality Report To Excel
        [HttpGet]
        [Route("ExportProjectLevelQualityReportToExcel")]
        public HttpResponseMessage ExportProjectLevelQualityReportToExcel(string CustomerCode, string ProjectCode, string BatchNo = "", string UserID = "", string Status = "P")
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Project Level Quality Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    //Swagger fix
                    if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                        BatchNo = "";

                    DataTable dataTable = new DataTable();
                    dataTable = _BLLProduction.ReadProjectLevelQualityReportSKUDetails(CustomerCode, ProjectCode, BatchNo, Status);

                    if (dataTable.Rows.Count > 0)
                    {
                        if (HttpContext.Current == null)
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);

                        //Create HTTP Response.
                        HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                        string FileName = "Project Level Quality Report.xlsx";

                        #region Setting up the workbook
                        AsposeHelpers asposeHelpers = new AsposeHelpers();
                        Workbook wb = asposeHelpers.GetWorkbook();
                        var ws = wb.Worksheets[0];
                        int row = 1;
                        #endregion

                        #region Setting Styles
                        Aspose.Cells.Style styleHeader = asposeHelpers.GetStyle(wb, 0, "header");
                        Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");
                        Aspose.Cells.Style styleLeftAlignData = asposeHelpers.GetStyle(wb, 0, "left");
                        #endregion

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("UniqueID");
                        ws.Cells[0, 2].PutValue("Short Description");
                        ws.Cells[0, 3].PutValue("Long Description");
                        ws.Cells[0, 4].PutValue("Production User");
                        ws.Cells[0, 5].PutValue("Production Status");
                        ws.Cells[0, 6].PutValue("QC User");
                        ws.Cells[0, 7].PutValue("QC Status");
                        ws.Cells[0, 8].PutValue("UOM");
                        ws.Cells[0, 9].PutValue("MFR Name");
                        ws.Cells[0, 10].PutValue("MFR P/N");
                        ws.Cells[0, 11].PutValue("Vendor Name");
                        ws.Cells[0, 12].PutValue("Vendor P/N");
                        ws.Cells[0, 13].PutValue("New Short Description");
                        ws.Cells[0, 14].PutValue("New Long Description");
                        ws.Cells[0, 15].PutValue("Missing Words");
                        ws.Cells[0, 16].PutValue("Noun");
                        ws.Cells[0, 17].PutValue("Modifier");
                        ws.Cells[0, 18].PutValue("Level");
                        ws.Cells[0, 19].PutValue("MFR Name1");
                        ws.Cells[0, 20].PutValue("MFR PN1");
                        ws.Cells[0, 21].PutValue("MFR Name2");
                        ws.Cells[0, 22].PutValue("MFR PN2");
                        ws.Cells[0, 23].PutValue("MFR Name3");
                        ws.Cells[0, 24].PutValue("MFR PN3");
                        ws.Cells[0, 25].PutValue("Vendor Name1");
                        ws.Cells[0, 26].PutValue("Vendor PN1");
                        ws.Cells[0, 27].PutValue("Vendor Name2");
                        ws.Cells[0, 28].PutValue("Vendor PN2");
                        ws.Cells[0, 29].PutValue("Vendor Name3");
                        ws.Cells[0, 30].PutValue("Vendor PN3");
                        ws.Cells[0, 31].PutValue("Additional Info");
                        ws.Cells[0, 32].PutValue("Additional Info From Web");
                        ws.Cells[0, 33].PutValue("UNSPSC Code");
                        ws.Cells[0, 34].PutValue("UNSPSC Category");
                        ws.Cells[0, 35].PutValue("Web Ref URL1");
                        ws.Cells[0, 36].PutValue("Web Ref URL2");
                        ws.Cells[0, 37].PutValue("Web Ref URL3");
                        ws.Cells[0, 38].PutValue("PDF URL");
                        ws.Cells[0, 39].PutValue("Remarks");
                        ws.Cells[0, 40].PutValue("Query");
                        ws.Cells[0, 41].PutValue("Application");
                        ws.Cells[0, 42].PutValue("DWG");
                        ws.Cells[0, 43].PutValue("POS");
                        ws.Cells[0, 44].PutValue("ItemNo");
                        ws.Cells[0, 45].PutValue("SerialNo");
                        ws.Cells[0, 46].PutValue("OtherNo");
                        ws.Cells[0, 47].PutValue("KKSCode");
                        ws.Cells[0, 48].PutValue("Assembly / Part");
                        ws.Cells[0, 49].PutValue("BOM");
                        ws.Cells[0, 50].PutValue("GreenItems");

                        for (int c = 0; c <= 50; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(dr["UniqueID"].ToString().Trim());
                            ws.Cells[row, 2].PutValue(dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 3].PutValue(dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 4].PutValue(dr["ProductionUser"].ToString());
                            ws.Cells[row, 5].PutValue(dr["ProductionStatus"].ToString());
                            ws.Cells[row, 6].PutValue(dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "");
                            ws.Cells[row, 7].PutValue(dr["QCStatus"] != DBNull.Value ? dr["QCStatus"].ToString().Trim() : "");
                            ws.Cells[row, 8].PutValue(dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "");
                            ws.Cells[row, 9].PutValue(dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "");
                            ws.Cells[row, 10].PutValue(dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "");
                            ws.Cells[row, 11].PutValue(dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "");
                            ws.Cells[row, 12].PutValue(dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "");
                            ws.Cells[row, 13].PutValue(dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 14].PutValue(dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 15].PutValue(dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "");
                            ws.Cells[row, 16].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                            ws.Cells[row, 17].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                            ws.Cells[row, 18].PutValue(dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "");
                            ws.Cells[row, 19].PutValue(dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "");
                            ws.Cells[row, 20].PutValue(dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "");
                            ws.Cells[row, 21].PutValue(dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "");
                            ws.Cells[row, 22].PutValue(dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "");
                            ws.Cells[row, 23].PutValue(dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "");
                            ws.Cells[row, 24].PutValue(dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "");
                            ws.Cells[row, 25].PutValue(dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "");
                            ws.Cells[row, 26].PutValue(dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "");
                            ws.Cells[row, 27].PutValue(dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "");
                            ws.Cells[row, 28].PutValue(dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "");
                            ws.Cells[row, 29].PutValue(dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "");
                            ws.Cells[row, 30].PutValue(dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "");
                            ws.Cells[row, 31].PutValue(dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "");
                            ws.Cells[row, 32].PutValue(dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "");
                            ws.Cells[row, 33].PutValue(dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "");
                            ws.Cells[row, 34].PutValue(dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "");
                            ws.Cells[row, 35].PutValue(dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "");
                            ws.Cells[row, 36].PutValue(dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "");
                            ws.Cells[row, 37].PutValue(dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "");
                            ws.Cells[row, 38].PutValue(dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "");
                            ws.Cells[row, 39].PutValue(dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "");
                            ws.Cells[row, 40].PutValue(dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "");
                            ws.Cells[row, 41].PutValue(dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "");
                            ws.Cells[row, 42].PutValue(dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "");
                            ws.Cells[row, 43].PutValue(dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "");
                            ws.Cells[row, 44].PutValue(dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "");
                            ws.Cells[row, 45].PutValue(dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "");
                            ws.Cells[row, 46].PutValue(dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "");
                            ws.Cells[row, 47].PutValue(dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "");
                            ws.Cells[row, 48].PutValue(dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "");
                            ws.Cells[row, 49].PutValue(dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "");
                            ws.Cells[row, 50].PutValue(dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "");
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
                            ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 12].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 14].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 15].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 25].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 27].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 28].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 29].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 31].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 32].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 33].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 35].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 36].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 37].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 38].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 39].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 40].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 41].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 42].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 47].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 48].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 49].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 50].SetStyle(styleCenterAlignData);
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
                    {
                        Result objResult = new Result();
                        objResult.Msg = "No data found";
                        objResult.Success = 0;

                        return Request.CreateResponse(HttpStatusCode.OK, objResult);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }

        }
        #endregion
        #endregion

        #region Resource Level Quality Report
        #region Resource Level Quality Report Count Stats
        [HttpGet]
        [Route("ReadResourceLevelQualityReportCountStats")]
        public HttpResponseMessage ReadResourceLevelQualityReportCountStats(string UserName, string CustomerCode, string ProjectCode, DateTime? FromDate, DateTime? ToDate, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadResourceLevelQualityCountStats(UserName, CustomerCode, ProjectCode, BatchNo, FromDate, ToDate);

                if (dataTable.Rows.Count > 0)
                {
                    JObject PEReport = new JObject(new JProperty("AllocatedCount", dataTable.Rows[0]["AllocatedCount"]),
                                                   new JProperty("ProcessedCount", dataTable.Rows[0]["ProcessedCount"]),
                                                   new JProperty("QCApprovedCount", dataTable.Rows[0]["QCApprovedCount"]),
                                                   new JProperty("AttributeChangedCount", dataTable.Rows[0]["AttributeChangedCount"]),
                                                   new JProperty("AttributeChangedPercentage", dataTable.Rows[0]["AttributeChangedPercentage"]),
                                                   new JProperty("NMChangedCount", dataTable.Rows[0]["NMChangedCount"]),
                                                   new JProperty("NMChangedPercentage", dataTable.Rows[0]["NMChangedPercentage"]),
                                                   new JProperty("MFRorSupplierChangedCount", dataTable.Rows[0]["MFRorSupplierChangedCount"]),
                                                   new JProperty("MFRorSupplierChangedPercentage", dataTable.Rows[0]["MFRorSupplierChangedPercentage"]),
                                                   new JProperty("ProjectScope", dataTable.Rows[0]["ProjectScope"]));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Resource Level Quality Report SKU Details
        [HttpGet]
        [Route("ReadResourceLevelQualityReportSKUDetails")]
        public HttpResponseMessage ReadResourceLevelQualityReportSKUDetails(string UserName, string CustomerCode, string ProjectCode, DateTime? FromDate, DateTime? ToDate, string BatchNo = "", string Status = "P")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadResourceLevelQualityReportSKUDetails(UserName, CustomerCode, ProjectCode, BatchNo, FromDate, ToDate, Status);

                List<ProductionItem> productionItemsList = new List<ProductionItem>();
                if (dataTable.Rows.Count > 0)
                {
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        ProductionItem productionItem = new ProductionItem();
                        productionItem.CustomerCode = dr["CustomerCode"].ToString();
                        productionItem.ProjectCode = dr["ProjectCode"].ToString();
                        productionItem.BatchNo = dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "";
                        productionItem.UniqueID = dr["UniqueID"].ToString().Trim();
                        productionItem.ShortDescription = dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "";
                        productionItem.LongDescription = dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "";
                        productionItem.ProductionUser = dr["ProductionUser"].ToString();
                        productionItem.Status = dr["ProductionStatus"].ToString();
                        productionItem.QCUser = dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "";
                        productionItem.QCStatus = dr["QCStatus"] != DBNull.Value ? dr["QCStatus"].ToString().Trim() : "";
                        productionItem.UOM = dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "";
                        productionItem.MFRName = dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "";
                        productionItem.MFRPN = dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "";
                        productionItem.VendorName = dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "";
                        productionItem.VendorPN = dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "";
                        productionItem.NewShortDescription = dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "";
                        productionItem.NewLongDescription = dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "";
                        productionItem.MissingWords = dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "";
                        productionItem.Noun = dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "";
                        productionItem.Modifier = dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "";
                        productionItem.Level = dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "";
                        productionItem.MFRName1 = dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "";
                        productionItem.MFRPN1 = dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "";
                        productionItem.MFRName2 = dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "";
                        productionItem.MFRPN2 = dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "";
                        productionItem.MFRName3 = dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "";
                        productionItem.MFRPN3 = dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "";
                        productionItem.VendorName1 = dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "";
                        productionItem.VendorPN1 = dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "";
                        productionItem.VendorName2 = dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "";
                        productionItem.VendorPN2 = dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "";
                        productionItem.VendorName3 = dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "";
                        productionItem.VendorPN3 = dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "";
                        productionItem.AdditionalInfo = dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "";
                        productionItem.AdditionalInfoFromWeb = dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "";
                        productionItem.UNSPSCCode = dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "";
                        productionItem.UNSPSCCategory = dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "";
                        productionItem.WebRefURL1 = dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "";
                        productionItem.WebRefURL2 = dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "";
                        productionItem.WebRefURL3 = dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "";
                        productionItem.PDFURL = dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "";
                        productionItem.Remarks = dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "";
                        productionItem.Query = dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "";
                        productionItem.Application = dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "";
                        productionItem.DWG = dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "";
                        productionItem.POS = dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "";
                        productionItem.ItemNo = dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "";
                        productionItem.SerialNo = dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "";
                        productionItem.OtherNo = dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "";
                        productionItem.KKSCode = dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "";
                        productionItem.AssemblyOrPart = dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "";
                        productionItem.BOM = dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "";
                        productionItem.GreenItems = dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "";
                        productionItemsList.Add(productionItem);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", productionItemsList.Count),
                    new JProperty("ProductionItemsList", new JArray(from i in productionItemsList
                                                                    select new JObject(
                                                                new JProperty("CustomerCode", i.CustomerCode),
                                                                new JProperty("ProjectCode", i.ProjectCode),
                                                                new JProperty("BatchNo", i.BatchNo),
                                                                new JProperty("UniqueID", i.UniqueID),
                                                                new JProperty("ShortDescription", i.ShortDescription),
                                                                new JProperty("LongDescription", i.LongDescription),
                                                                new JProperty("ProductionUser", i.ProductionUser),
                                                                new JProperty("ProductionStatus", i.Status),
                                                                new JProperty("QCUser", i.QCUser),
                                                                new JProperty("QCStatus", i.QCStatus),
                                                                new JProperty("MFRName", i.MFRName),
                                                                new JProperty("MFRPN", i.MFRPN),
                                                                new JProperty("VendorName", i.VendorName),
                                                                new JProperty("VendorPN", i.VendorPN),
                                                                new JProperty("UOM", i.UOM),
                                                                new JProperty("NewShortDescription", i.NewShortDescription),
                                                                new JProperty("NewLongDescription", i.NewLongDescription),
                                                                new JProperty("MissingWords", i.MissingWords),
                                                                new JProperty("Noun", i.Noun),
                                                                new JProperty("Modifier", i.Modifier),
                                                                new JProperty("Level", i.Level),
                                                                new JProperty("MFRName1", i.MFRName1),
                                                                new JProperty("MFRPN1", i.MFRPN1),
                                                                new JProperty("MFRName2", i.MFRName2),
                                                                new JProperty("MFRPN2", i.MFRPN2),
                                                                new JProperty("MFRName3", i.MFRName3),
                                                                new JProperty("MFRPN3", i.MFRPN3),
                                                                new JProperty("VendorName1", i.VendorName1),
                                                                new JProperty("VendorPN1", i.VendorPN1),
                                                                new JProperty("VendorName2", i.VendorName2),
                                                                new JProperty("VendorPN2", i.VendorPN2),
                                                                new JProperty("VendorName3", i.VendorName3),
                                                                new JProperty("VendorPN3", i.VendorPN3),
                                                                new JProperty("AdditionalInfoFromInput", i.AdditionalInfo),
                                                                new JProperty("AdditionalInfoFromWeb", i.AdditionalInfoFromWeb),
                                                                new JProperty("UNSPSCCode", i.UNSPSCCode),
                                                                new JProperty("UNSPSCCategory", i.UNSPSCCategory),
                                                                new JProperty("WebRefURL1", i.WebRefURL1),
                                                                new JProperty("WebRefURL2", i.WebRefURL2),
                                                                new JProperty("WebRefURL3", i.WebRefURL3),
                                                                new JProperty("PDFURL", i.PDFURL),
                                                                new JProperty("Remarks", i.Remarks),
                                                                new JProperty("Query", i.Query),
                                                                new JProperty("Application", i.Application),
                                                                new JProperty("DWG", i.DWG),
                                                                new JProperty("ItemNo", i.ItemNo),
                                                                new JProperty("POS", i.POS),
                                                                new JProperty("SerialNo", i.SerialNo),
                                                                new JProperty("OtherNo", i.OtherNo),
                                                                new JProperty("KKSCode", i.KKSCode),
                                                                new JProperty("AssemblyOrPart", i.AssemblyOrPart),
                                                                new JProperty("BOM", i.BOM),
                                                                new JProperty("GreenItems", i.GreenItems)
                                                                ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export Resource Level Quality Report To Excel
        [HttpGet]
        [Route("ExportResourceLevelQualityReportToExcel")]
        public HttpResponseMessage ExportResourceLevelQualityReportToExcel(string UserName, string CustomerCode, string ProjectCode, DateTime? FromDate, DateTime? ToDate, string BatchNo = "", string UserID = "", string Status = "P")
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Project Level Quality Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    //Swagger fix
                    if (BatchNo == "{BatchNo}" || string.IsNullOrEmpty(BatchNo))
                        BatchNo = "";

                    DataTable dataTable = new DataTable();
                    dataTable = _BLLProduction.ReadResourceLevelQualityReportSKUDetails(UserName, CustomerCode, ProjectCode, BatchNo, FromDate, ToDate, Status);

                    if (dataTable.Rows.Count > 0)
                    {
                        if (HttpContext.Current == null)
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);

                        //Create HTTP Response.
                        HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                        string FileName = "Resource Level Quality Report.xlsx";

                        #region Setting up the workbook
                        AsposeHelpers asposeHelpers = new AsposeHelpers();
                        Workbook wb = asposeHelpers.GetWorkbook();
                        var ws = wb.Worksheets[0];
                        int row = 1;
                        #endregion

                        #region Setting Styles
                        Aspose.Cells.Style styleHeader = asposeHelpers.GetStyle(wb, 0, "header");
                        Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");
                        Aspose.Cells.Style styleLeftAlignData = asposeHelpers.GetStyle(wb, 0, "left");
                        #endregion

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("UniqueID");
                        ws.Cells[0, 2].PutValue("Short Description");
                        ws.Cells[0, 3].PutValue("Long Description");
                        ws.Cells[0, 4].PutValue("Production User");
                        ws.Cells[0, 5].PutValue("Production Status");
                        ws.Cells[0, 6].PutValue("QC User");
                        ws.Cells[0, 7].PutValue("QC Status");
                        ws.Cells[0, 8].PutValue("UOM");
                        ws.Cells[0, 9].PutValue("MFR Name");
                        ws.Cells[0, 10].PutValue("MFR P/N");
                        ws.Cells[0, 11].PutValue("Vendor Name");
                        ws.Cells[0, 12].PutValue("Vendor P/N");
                        ws.Cells[0, 13].PutValue("New Short Description");
                        ws.Cells[0, 14].PutValue("New Long Description");
                        ws.Cells[0, 15].PutValue("Missing Words");
                        ws.Cells[0, 16].PutValue("Noun");
                        ws.Cells[0, 17].PutValue("Modifier");
                        ws.Cells[0, 18].PutValue("Level");
                        ws.Cells[0, 19].PutValue("MFR Name1");
                        ws.Cells[0, 20].PutValue("MFR PN1");
                        ws.Cells[0, 21].PutValue("MFR Name2");
                        ws.Cells[0, 22].PutValue("MFR PN2");
                        ws.Cells[0, 23].PutValue("MFR Name3");
                        ws.Cells[0, 24].PutValue("MFR PN3");
                        ws.Cells[0, 25].PutValue("Vendor Name1");
                        ws.Cells[0, 26].PutValue("Vendor PN1");
                        ws.Cells[0, 27].PutValue("Vendor Name2");
                        ws.Cells[0, 28].PutValue("Vendor PN2");
                        ws.Cells[0, 29].PutValue("Vendor Name3");
                        ws.Cells[0, 30].PutValue("Vendor PN3");
                        ws.Cells[0, 31].PutValue("Additional Info");
                        ws.Cells[0, 32].PutValue("Additional Info From Web");
                        ws.Cells[0, 33].PutValue("UNSPSC Code");
                        ws.Cells[0, 34].PutValue("UNSPSC Category");
                        ws.Cells[0, 35].PutValue("Web Ref URL1");
                        ws.Cells[0, 36].PutValue("Web Ref URL2");
                        ws.Cells[0, 37].PutValue("Web Ref URL3");
                        ws.Cells[0, 38].PutValue("PDF URL");
                        ws.Cells[0, 39].PutValue("Remarks");
                        ws.Cells[0, 40].PutValue("Query");
                        ws.Cells[0, 41].PutValue("Application");
                        ws.Cells[0, 42].PutValue("DWG");
                        ws.Cells[0, 43].PutValue("POS");
                        ws.Cells[0, 44].PutValue("ItemNo");
                        ws.Cells[0, 45].PutValue("SerialNo");
                        ws.Cells[0, 46].PutValue("OtherNo");
                        ws.Cells[0, 47].PutValue("KKSCode");
                        ws.Cells[0, 48].PutValue("Assembly / Part");
                        ws.Cells[0, 49].PutValue("BOM");
                        ws.Cells[0, 50].PutValue("GreenItems");
                        ws.Cells[0, 51].PutValue("Customer Code");
                        ws.Cells[0, 52].PutValue("Project Code");
                        ws.Cells[0, 53].PutValue("Batch No.");

                        for (int c = 0; c <= 53; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(dr["UniqueID"].ToString().Trim());
                            ws.Cells[row, 2].PutValue(dr["ShortDescription"] != DBNull.Value ? dr["ShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 3].PutValue(dr["LongDescription"] != DBNull.Value ? dr["LongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 4].PutValue(dr["ProductionUser"].ToString());
                            ws.Cells[row, 5].PutValue(dr["ProductionStatus"].ToString());
                            ws.Cells[row, 6].PutValue(dr["QCUser"] != DBNull.Value ? dr["QCUser"].ToString().Trim() : "");
                            ws.Cells[row, 7].PutValue(dr["QCStatus"] != DBNull.Value ? dr["QCStatus"].ToString().Trim() : "");
                            ws.Cells[row, 8].PutValue(dr["UOM"] != DBNull.Value ? dr["UOM"].ToString().Trim() : "");
                            ws.Cells[row, 9].PutValue(dr["MFRName"] != DBNull.Value ? dr["MFRName"].ToString().Trim() : "");
                            ws.Cells[row, 10].PutValue(dr["MFRPN"] != DBNull.Value ? dr["MFRPN"].ToString().Trim() : "");
                            ws.Cells[row, 11].PutValue(dr["VendorName"] != DBNull.Value ? dr["VendorName"].ToString().Trim() : "");
                            ws.Cells[row, 12].PutValue(dr["VendorPN"] != DBNull.Value ? dr["VendorPN"].ToString().Trim() : "");
                            ws.Cells[row, 13].PutValue(dr["NewShortDescription"] != DBNull.Value ? dr["NewShortDescription"].ToString().Trim() : "");
                            ws.Cells[row, 14].PutValue(dr["NewLongDescription"] != DBNull.Value ? dr["NewLongDescription"].ToString().Trim() : "");
                            ws.Cells[row, 15].PutValue(dr["MissingWords"] != DBNull.Value ? dr["MissingWords"].ToString().Trim() : "");
                            ws.Cells[row, 16].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                            ws.Cells[row, 17].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                            ws.Cells[row, 18].PutValue(dr["Level"] != DBNull.Value ? dr["Level"].ToString().Trim() : "");
                            ws.Cells[row, 19].PutValue(dr["MFRName1"] != DBNull.Value ? dr["MFRName1"].ToString().Trim() : "");
                            ws.Cells[row, 20].PutValue(dr["MFRPN1"] != DBNull.Value ? dr["MFRPN1"].ToString().Trim() : "");
                            ws.Cells[row, 21].PutValue(dr["MFRName2"] != DBNull.Value ? dr["MFRName2"].ToString().Trim() : "");
                            ws.Cells[row, 22].PutValue(dr["MFRPN2"] != DBNull.Value ? dr["MFRPN2"].ToString().Trim() : "");
                            ws.Cells[row, 23].PutValue(dr["MFRName3"] != DBNull.Value ? dr["MFRName3"].ToString().Trim() : "");
                            ws.Cells[row, 24].PutValue(dr["MFRPN3"] != DBNull.Value ? dr["MFRPN3"].ToString().Trim() : "");
                            ws.Cells[row, 25].PutValue(dr["VendorName1"] != DBNull.Value ? dr["VendorName1"].ToString().Trim() : "");
                            ws.Cells[row, 26].PutValue(dr["VendorPN1"] != DBNull.Value ? dr["VendorPN1"].ToString().Trim() : "");
                            ws.Cells[row, 27].PutValue(dr["VendorName2"] != DBNull.Value ? dr["VendorName2"].ToString().Trim() : "");
                            ws.Cells[row, 28].PutValue(dr["VendorPN2"] != DBNull.Value ? dr["VendorPN2"].ToString().Trim() : "");
                            ws.Cells[row, 29].PutValue(dr["VendorName3"] != DBNull.Value ? dr["VendorName3"].ToString().Trim() : "");
                            ws.Cells[row, 30].PutValue(dr["VendorPN3"] != DBNull.Value ? dr["VendorPN3"].ToString().Trim() : "");
                            ws.Cells[row, 31].PutValue(dr["AdditionalInfoFromInput"] != DBNull.Value ? dr["AdditionalInfoFromInput"].ToString().Trim() : "");
                            ws.Cells[row, 32].PutValue(dr["AdditionalInfoFromWeb"] != DBNull.Value ? dr["AdditionalInfoFromWeb"].ToString().Trim() : "");
                            ws.Cells[row, 33].PutValue(dr["UNSPSCCode"] != DBNull.Value ? dr["UNSPSCCode"].ToString().Trim() : "");
                            ws.Cells[row, 34].PutValue(dr["UNSPSCCategory"] != DBNull.Value ? dr["UNSPSCCategory"].ToString().Trim() : "");
                            ws.Cells[row, 35].PutValue(dr["WebRefURL1"] != DBNull.Value ? dr["WebRefURL1"].ToString().Trim() : "");
                            ws.Cells[row, 36].PutValue(dr["WebRefURL2"] != DBNull.Value ? dr["WebRefURL2"].ToString().Trim() : "");
                            ws.Cells[row, 37].PutValue(dr["WebRefURL3"] != DBNull.Value ? dr["WebRefURL3"].ToString().Trim() : "");
                            ws.Cells[row, 38].PutValue(dr["PDFURL"] != DBNull.Value ? dr["PDFURL"].ToString().Trim() : "");
                            ws.Cells[row, 39].PutValue(dr["Remarks"] != DBNull.Value ? dr["Remarks"].ToString().Trim() : "");
                            ws.Cells[row, 40].PutValue(dr["Query"] != DBNull.Value ? dr["Query"].ToString().Trim() : "");
                            ws.Cells[row, 41].PutValue(dr["Application"] != DBNull.Value ? dr["Application"].ToString().Trim() : "");
                            ws.Cells[row, 42].PutValue(dr["DWG"] != DBNull.Value ? dr["DWG"].ToString().Trim() : "");
                            ws.Cells[row, 43].PutValue(dr["POS"] != DBNull.Value ? dr["POS"].ToString().Trim() : "");
                            ws.Cells[row, 44].PutValue(dr["ItemNo"] != DBNull.Value ? dr["ItemNo"].ToString().Trim() : "");
                            ws.Cells[row, 45].PutValue(dr["SerialNo"] != DBNull.Value ? dr["SerialNo"].ToString().Trim() : "");
                            ws.Cells[row, 46].PutValue(dr["OtherNo"] != DBNull.Value ? dr["OtherNo"].ToString().Trim() : "");
                            ws.Cells[row, 47].PutValue(dr["KKSCode"] != DBNull.Value ? dr["KKSCode"].ToString().Trim() : "");
                            ws.Cells[row, 48].PutValue(dr["AssemblyOrPart"] != DBNull.Value ? dr["AssemblyOrPart"].ToString().Trim() : "");
                            ws.Cells[row, 49].PutValue(dr["BOM"] != DBNull.Value ? dr["BOM"].ToString().Trim() : "");
                            ws.Cells[row, 50].PutValue(dr["GreenItems"] != DBNull.Value ? dr["GreenItems"].ToString().Trim() : "");
                            ws.Cells[row, 51].PutValue(dr["CustomerCode"].ToString());
                            ws.Cells[row, 52].PutValue(dr["ProjectCode"].ToString());
                            ws.Cells[row, 53].PutValue(dr["BatchNo"] != DBNull.Value ? dr["BatchNo"].ToString().Trim() : "");
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
                            ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 12].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 14].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 15].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 24].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 25].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 26].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 27].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 28].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 29].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 30].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 31].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 32].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 33].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 34].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 35].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 36].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 37].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 38].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 39].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 40].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 41].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 42].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 43].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 44].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 45].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 46].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 47].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 48].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 49].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 50].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 51].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 52].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 53].SetStyle(styleCenterAlignData);
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
                    {
                        Result objResult = new Result();
                        objResult.Msg = "No data found";
                        objResult.Success = 0;

                        return Request.CreateResponse(HttpStatusCode.OK, objResult);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }

        }
        #endregion
        #endregion

        #region Read MRO Ref. DB Parts List by Part No.
        [HttpGet]
        [Route("ReadMRORefDBPartDetailsByPartNo")]
        public IHttpActionResult ReadMRORefDBPartDetailsByPartNo(string PartNo)
        {
            try
            {
                List<MRORefDBPartDetails> mroRefDBPartDetailsList = new List<MRORefDBPartDetails>();
                DataTable dataTable = new DataTable();
                dataTable = _BLLProduction.ReadMRORefDBPartDetailsByPartNo(PartNo);
                if (dataTable.Rows.Count > 0)
                {
                    MRORefDBPartDetails partDetails = new MRORefDBPartDetails();
                    DataRow dr = dataTable.Rows[0];
                    partDetails.MRORefDBID = Convert.ToInt64(dr["MRORefDBID"]);
                    partDetails.MFRName1 = dr["MANUFACTURER NAME"] != DBNull.Value ? dr["MANUFACTURER NAME"].ToString().Trim() : "";
                    partDetails.MFRPN1 = dr["MANUFACTURER PART NO"] != DBNull.Value ? dr["MANUFACTURER PART NO"].ToString().Trim() : "";
                    partDetails.MFRName2 = dr["MANUFACTURER NAME_1"] != DBNull.Value ? dr["MANUFACTURER NAME_1"].ToString().Trim() : "";
                    partDetails.MFRPN2 = dr["MANUFACTURER PART NO_1"] != DBNull.Value ? dr["MANUFACTURER PART NO_1"].ToString().Trim() : "";
                    partDetails.RefURL1 = dr["Reference URL"] != DBNull.Value ? dr["Reference URL"].ToString().Trim() : "";
                    partDetails.RefURL2 = dr["Reference URL 1"] != DBNull.Value ? dr["Reference URL 1"].ToString().Trim() : "";
                    partDetails.Noun = dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "";
                    partDetails.Modifier = dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "";
                    partDetails.AttributeName1 = dr["Attribute Name1"] != DBNull.Value ? dr["Attribute Name1"].ToString().Trim() : "";
                    partDetails.AttributeValue1 = dr["Attribute Value1"] != DBNull.Value ? dr["Attribute Value1"].ToString().Trim() : "";
                    partDetails.AttributeName2 = dr["Attribute Name2"] != DBNull.Value ? dr["Attribute Name2"].ToString().Trim() : "";
                    partDetails.AttributeValue2 = dr["Attribute Value2"] != DBNull.Value ? dr["Attribute Value2"].ToString().Trim() : "";
                    partDetails.AttributeName3 = dr["Attribute Name3"] != DBNull.Value ? dr["Attribute Name3"].ToString().Trim() : "";
                    partDetails.AttributeValue3 = dr["Attribute Value3"] != DBNull.Value ? dr["Attribute Value3"].ToString().Trim() : "";
                    partDetails.AttributeName4 = dr["Attribute Name4"] != DBNull.Value ? dr["Attribute Name4"].ToString().Trim() : "";
                    partDetails.AttributeValue4 = dr["Attribute Value4"] != DBNull.Value ? dr["Attribute Value4"].ToString().Trim() : "";
                    partDetails.AttributeName5 = dr["Attribute Name5"] != DBNull.Value ? dr["Attribute Name5"].ToString().Trim() : "";
                    partDetails.AttributeValue5 = dr["Attribute Value5"] != DBNull.Value ? dr["Attribute Value5"].ToString().Trim() : "";
                    partDetails.AttributeName6 = dr["Attribute Name6"] != DBNull.Value ? dr["Attribute Name6"].ToString().Trim() : "";
                    partDetails.AttributeValue6 = dr["Attribute Value6"] != DBNull.Value ? dr["Attribute Value6"].ToString().Trim() : "";
                    partDetails.AttributeName7 = dr["Attribute Name7"] != DBNull.Value ? dr["Attribute Name7"].ToString().Trim() : "";
                    partDetails.AttributeValue7 = dr["Attribute Value7"] != DBNull.Value ? dr["Attribute Value7"].ToString().Trim() : "";
                    partDetails.AttributeName8 = dr["Attribute Name8"] != DBNull.Value ? dr["Attribute Name8"].ToString().Trim() : "";
                    partDetails.AttributeValue8 = dr["Attribute Value8"] != DBNull.Value ? dr["Attribute Value8"].ToString().Trim() :
                    partDetails.AttributeName9 = dr["Attribute Name9"] != DBNull.Value ? dr["Attribute Name9"].ToString().Trim() : "";
                    partDetails.AttributeValue9 = dr["Attribute Value9"] != DBNull.Value ? dr["Attribute Value9"].ToString().Trim() : "";
                    partDetails.AttributeName10 = dr["Attribute Name10"] != DBNull.Value ? dr["Attribute Name10"].ToString().Trim() : "";
                    partDetails.AttributeValue10 = dr["Attribute Value10"] != DBNull.Value ? dr["Attribute Value10"].ToString().Trim() : "";
                    partDetails.AttributeName11 = dr["Attribute Name11"] != DBNull.Value ? dr["Attribute Name11"].ToString().Trim() : "";
                    partDetails.AttributeValue11 = dr["Attribute Value11"] != DBNull.Value ? dr["Attribute Value11"].ToString().Trim() : "";
                    partDetails.AttributeName12 = dr["Attribute Name12"] != DBNull.Value ? dr["Attribute Name12"].ToString().Trim() : "";
                    partDetails.AttributeValue12 = dr["Attribute Value12"] != DBNull.Value ? dr["Attribute Value12"].ToString().Trim() : "";
                    partDetails.AttributeName13 = dr["Attribute Name13"] != DBNull.Value ? dr["Attribute Name13"].ToString().Trim() : "";
                    partDetails.AttributeValue13 = dr["Attribute Value13"] != DBNull.Value ? dr["Attribute Value13"].ToString().Trim() : "";
                    partDetails.AttributeName14 = dr["Attribute Name14"] != DBNull.Value ? dr["Attribute Name14"].ToString().Trim() : "";
                    partDetails.AttributeValue14 = dr["Attribute Value14"] != DBNull.Value ? dr["Attribute Value14"].ToString().Trim() : "";
                    partDetails.AttributeName15 = dr["Attribute Name15"] != DBNull.Value ? dr["Attribute Name15"].ToString().Trim() : "";
                    partDetails.AttributeValue15 = dr["Attribute Value15"] != DBNull.Value ? dr["Attribute Value15"].ToString().Trim() : "";
                    partDetails.AttributeName16 = dr["Attribute Name16"] != DBNull.Value ? dr["Attribute Name16"].ToString().Trim() : "";
                    partDetails.AttributeValue16 = dr["Attribute Value16"] != DBNull.Value ? dr["Attribute Value16"].ToString().Trim() : "";
                    partDetails.AttributeName17 = dr["Attribute Name17"] != DBNull.Value ? dr["Attribute Name17"].ToString().Trim() : "";
                    partDetails.AttributeValue17 = dr["Attribute Value17"] != DBNull.Value ? dr["Attribute Value17"].ToString().Trim() : "";
                    partDetails.AttributeName18 = dr["Attribute Name18"] != DBNull.Value ? dr["Attribute Name18"].ToString().Trim() : "";
                    partDetails.AttributeValue18 = dr["Attribute Value18"] != DBNull.Value ? dr["Attribute Value18"].ToString().Trim() : "";
                    partDetails.AttributeName19 = dr["Attribute Name19"] != DBNull.Value ? dr["Attribute Name19"].ToString().Trim() : "";
                    partDetails.AttributeValue19 = dr["Attribute Value19"] != DBNull.Value ? dr["Attribute Value19"].ToString().Trim() : "";
                    partDetails.AttributeName20 = dr["Attribute Name20"] != DBNull.Value ? dr["Attribute Name20"].ToString().Trim() : "";
                    partDetails.AttributeValue20 = dr["Attribute Value20"] != DBNull.Value ? dr["Attribute Value20"].ToString().Trim() : "";
                    mroRefDBPartDetailsList.Add(partDetails);
                }

                return Ok(mroRefDBPartDetailsList);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read MRO Ref. DB Part Details by ID
        [HttpGet]
        [Route("MRORefDBPartDetailsByID")]
        public IHttpActionResult MRORefDBPartDetailsByID(long MRORefDBID)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                List<ItemAttribute> ItemAttributesList = new List<ItemAttribute>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    #region Part Attributes
                    SqlCommand cmdItemAttributes = new SqlCommand();
                    cmdItemAttributes.Connection = conn;
                    cmdItemAttributes.CommandType = CommandType.StoredProcedure;
                    cmdItemAttributes.CommandText = "spMRORefDBMFRPartAttributeDetails";
                    cmdItemAttributes.Parameters.AddWithValue("@MRORefDBID", MRORefDBID);

                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmdItemAttributes.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ItemAttribute ItemAttribute = new ItemAttribute();
                        ItemAttribute.AttributeName = sqlReader["AttributeName"].ToString();
                        ItemAttribute.AttributeValue = sqlReader["AttributeValue"].ToString();
                        ItemAttribute.QCAttributeValue = sqlReader["QCAttributeValue"].ToString();
                        ItemAttribute.QCAttributeValueComments = sqlReader["QCAttributeValueComments"].ToString();
                        ItemAttributesList.Add(ItemAttribute);
                    }
                    conn.Close();
                    #endregion

                    #region MRO Ref DB Part Details
                    MRORefDBPartDetails mroRefDBPartDetails = new MRORefDBPartDetails();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRORefDBMFRPartDetailsByID";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@MRORefDBID", MRORefDBID);

                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        mroRefDBPartDetails.MRORefDBID = Convert.ToInt64(sqlReader["MRORefDBID"]);
                        mroRefDBPartDetails.MFRName1 = sqlReader["MANUFACTURER NAME"].ToString().Trim();
                        mroRefDBPartDetails.MFRPN1 = sqlReader["MANUFACTURER PART NO"].ToString().Trim();
                        mroRefDBPartDetails.MFRName2 = sqlReader["MANUFACTURER NAME_1"].ToString().Trim();
                        mroRefDBPartDetails.MFRPN2 = sqlReader["MANUFACTURER PART NO_1"].ToString().Trim();
                        mroRefDBPartDetails.RefURL1 = sqlReader["Reference URL"].ToString().Trim();
                        mroRefDBPartDetails.RefURL2 = sqlReader["Reference URL 1"].ToString().Trim();
                        mroRefDBPartDetails.Noun = sqlReader["Noun"].ToString().Trim();
                        mroRefDBPartDetails.Modifier = sqlReader["Modifier"].ToString().Trim();
                        mroRefDBPartDetails.ItemAttributes = ItemAttributesList;
                    }
                    conn.Close();
                    #endregion

                    return Ok(mroRefDBPartDetails);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
    }
}

