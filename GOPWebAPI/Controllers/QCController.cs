using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Xml;
using System.Xml.Serialization;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.Incident_Report_Models;
using Newtonsoft.Json.Linq;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/QC")]
    public class QCController : ApiController
    {
        private BLLQC _BLLQC;
        public QCController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLQC = new BLLQC(connectionString);
        }

        #region Read Projects Customer Codes of User
        [HttpGet]
        [Route("ReadProjectCustomerCodesOfUser/{QCUser}/{status?}")]
        public IHttpActionResult ReadProjectCustomerCodesOfUser(string QCUser, string Status = "O")
        {
            try
            {
                List<Customer> CustomerCodeList = new List<Customer>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCGetCustomerCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Status", Status);
                    cmd.Parameters.AddWithValue("@QCUser", QCUser);

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
        [Route("ReadCustomerProjectCodesOfUser/{CustomerCode}/{QCUser}/{status?}")]
        public IHttpActionResult ReadCustomerProjectCodesOfUser(string CustomerCode, string QCUser, string Status = "O")
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> ProjectCodeList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCGetProjectCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@QCUser", QCUser);
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
        [Route("ReadCustomerProjectBatchNosOfUser/{CustomerCode}/{ProjectCode}/{QCUser}/{status?}")]
        public IHttpActionResult ReadCustomerProjectBatchNosOfUser(string CustomerCode, string ProjectCode, string QCUser, string Status = "O")
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> BatchNoList = new List<CustomerCodeProjectCodeBatchNo>();
                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCGetBatchNos";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@QCUser", QCUser);
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

        #region Read Project Activities of QC User
        [HttpGet]
        [Route("ReadProjectActivitiesOfUser/{CustomerCode}/{ProjectCode}/{QCUser}/{BatchNo?}")]
        public IHttpActionResult ReadProjectActivitiesOfUser(string CustomerCode, string ProjectCode, string QCUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Allocation Counts
                List<QCProjectActivitiesCount> ProjectActivitiesCountList = new List<QCProjectActivitiesCount>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCProjectActivityCountStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@QCUser", QCUser);

                    //Calling sp to get list of Allocation Counts
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCProjectActivitiesCount pac = new QCProjectActivitiesCount();

                        pac.QCAllocationID = Convert.ToInt64(sqlReader["QCAllocationID"]);
                        pac.Activities = sqlReader["Activities"].ToString();
                        pac.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                        pac.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        pac.QCPendingCount = Convert.ToInt32(sqlReader["QCPendingCount"]);
                        pac.ProductionErrorCount = Convert.ToInt32(sqlReader["ProductionErrorCount"]);
                        pac.IsAllocationDownloadedForQC = Convert.ToInt32(sqlReader["IsAllocationDownloadedForQC"]);
                        pac.IsProductionErrorUploadCompleted = Convert.ToInt32(sqlReader["IsProductionErrorUploadCompleted"]);

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

        #region QC Download
        #region Download QC Allocation of User
        [HttpPost]
        [Route("DownloadQCAllocationOfUser")]
        public HttpResponseMessage DownloadQCAllocationOfUser([FromBody] QCDownload qcDownload)
        {
            try
            {
                string FileName = "QCAllocation_" + qcDownload.QCUser + '_' + qcDownload.QCAllocationID.ToString() + ".xlsx";

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCUpdateDownload";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@QCAllocationID", qcDownload.QCAllocationID);
                    cmd.Parameters.AddWithValue("@UserID", qcDownload.QCUser);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "updated")
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/"));
                        string AllocatedFileName = Path.Combine(dirUploads.FullName, arrResult[1].ToString());
                        int QCUserColumnNo = -1, StatusColNo = -1, dRowNo = 0;

                        if (!File.Exists(AllocatedFileName))
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocation file not found");

                        #region Copy the Project Allocated User data from allocated file to new file and download the file
                        Workbook wbAF = new Workbook();         //Allocated file
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wbAF.LoadData(AllocatedFileName);
                        var ws = wbAF.Worksheets[0];            //source worksheet
                        //Cells cells = ws.Cells;

                        Workbook wbUDF = new Workbook();         //User Download file
                        Aspose.Cells.License l1 = new Aspose.Cells.License();
                        l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wbUDF.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                        var wd = wbUDF.Worksheets[0];           //destination worksheet

                        #region Copying header as is from allocated file
                        for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                        {
                            wd.Cells[0, sCol].Copy(ws.Cells[0, sCol]);
                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "qc user")
                                QCUserColumnNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "status")
                                StatusColNo = sCol;

                            if (QCUserColumnNo >= 0 && StatusColNo >= 0)
                                break;
                        }
                        #endregion

                        #region Copying data rows of user from allocated file
                        dRowNo = 1;
                        for (int sRow = 1; sRow <= ws.Cells.MaxRow; sRow++)
                        {
                            if (ws.Cells[sRow, QCUserColumnNo].StringValue.Trim().ToLower() == qcDownload.QCUser.Trim().ToLower())
                            {
                                //wd.Cells.CopyRow(cells, sRow, dRowNo);
                                for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                                    wd.Cells[dRowNo, sCol].Copy(ws.Cells[sRow, sCol]);
                                dRowNo++;
                            }
                        }
                        #endregion

                        #region Write Status as 'QC Downloaded' for all rows and adding 'QC Comments' column just heading
                        int QCCommentsColNo = wd.Cells.MaxColumn + 1;
                        wd.Cells[0, QCCommentsColNo].PutValue("QC Comments");
                        for (int dRow = 1; dRow <= wd.Cells.MaxRow; dRow++)
                            wd.Cells[dRow, StatusColNo].PutValue("QC Downloaded");
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

        #region QC Upload
        [HttpPost]
        [Route("ValidateAndUploadQC")]
        public HttpResponseMessage ValidateAndUploadQC([FromBody] QCUpload qcUpload)
        {
            try
            {
                string AllocatedFileName = string.Empty;
                string AllocatedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/");
                string AllocatedFileFullName = string.Empty;
                System.Data.Common.DbDataReader sqlReader;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid QC Upload Data");

                if (!AccessControl.CanUserAccessPage(qcUpload.UserID, "QC Download-Upload"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + qcUpload.UploadedFileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload QC Upload file");
                #endregion

                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + qcUpload.UploadedFileName);
                string UploadedFileExtension = Path.GetExtension(UploadedFilepath);
                int QCUserColumnNo=-1, ActivitiesColumnNo = -1, StatusColumnNo=-1;

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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded QC file has no data rows in first worksheet.");
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

                #region Check Activities, QC User, Status columns exists, if exists get column Nos.
                for (int col = 0; col <= wsUF.Cells.MaxColumn; col++)
                {
                    if (string.IsNullOrEmpty(wsUF.Cells[0, col].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Blank/Empty column heading in uploaded file. Column No.: " + (col + 1).ToString());

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "activities")
                        ActivitiesColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "qc user")
                        QCUserColumnNo = col;

                    if (wsUF.Cells[0, col].StringValue.Trim().ToLower() == "status")
                        StatusColumnNo = col;

                    if (ActivitiesColumnNo >= 0 && QCUserColumnNo >= 0 && StatusColumnNo>=0)
                        break;
                }

                if (ActivitiesColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column not found in uploaded QC file");

                if (QCUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'QC User' column not found in uploaded QC file");

                if (StatusColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Status' column not found in uploaded QC file");
                #endregion

                #region Check QC upload file should have QC allocated user entries only
                for (int row = 1; row <= wsUF.Cells.MaxRow; row++)
                {
                    if (wsUF.Cells[row, QCUserColumnNo].StringValue.Trim().ToLower() != qcUpload.UserID.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file should have only current QC user rows");

                    if (wsUF.Cells[row, StatusColumnNo].StringValue.Trim().ToLower() != "qc completed" && 
                        wsUF.Cells[row, StatusColumnNo].StringValue.Trim().ToLower() !="production error")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file SKU status can be only 'QC Completed' or 'Production Error'");
                }
                #endregion

                #region Get the allocated file name
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

                //Initialize command object
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "spQCAllocationGetUploadedFileNameByID";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                //Add parameters with values
                cmd.Parameters.AddWithValue("@QCAllocationID", qcUpload.QCAllocationID);

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

                //Check whether File exists
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
                    if (UploadedFileColumnName.Trim().ToLower() != "qc comments")
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, UploadedFileColumnName + " - column not found. Columns in uploaded QC file should be as same as allocated file");
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, AllocatedFileColumnName + " - column not found. Columns in uploaded QC file should be as same as allocated file");
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

                #region Create temp table, write uploaded file data, upload QC, Move file
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

                #region Create QC Upload
                //Initialize command object
                SqlCommand cmdQCU = new SqlCommand();
                cmdQCU.Connection = conn;
                cmdQCU.CommandType = CommandType.StoredProcedure;
                cmdQCU.CommandText = "spQCValidateAndUpload";

                #region Adding Stored Procedure Parameters
                cmdQCU.Parameters.AddWithValue("@QCAllocationID", qcUpload.QCAllocationID);
                cmdQCU.Parameters.AddWithValue("@UploadedFileExtension", UploadedFileExtension);
                cmdQCU.Parameters.AddWithValue("@TempTableName", sqlTableName);
                cmdQCU.Parameters.AddWithValue("@UserID", qcUpload.UserID);
                #endregion

                //Calling sp to create QC upload
                cmdQCU.CommandTimeout = 0;
                string Result = cmdQCU.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split(',');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    long QCUploadID = Convert.ToInt64(arrResult[1]);
                    string NewUploadedFileName = arrResult[2];

                    #region Input File move Starts
                    if (File.Exists(dirTemp + qcUpload.UploadedFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/QCUpload/"));
                        FileOperations.MoveFile(dirTemp, qcUpload.UploadedFileName, dirUploads, NewUploadedFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, QCUploadID);
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

        #region Production Error Download
        [HttpPost]
        [Route("DownloadProductionErrorSKUs")]
        public HttpResponseMessage DownloadProductionErrorSKUs([FromBody] QCDownload qcDownload)
        {
            try
            {
                string FileName = "ProductionErrorSKUs_" + qcDownload.QCUser + '_' + qcDownload.QCAllocationID.ToString() + ".xlsx";

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
                    cmd.CommandText = "spQCProductionErrorDownload";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@QCAllocationID", qcDownload.QCAllocationID);
                    cmd.Parameters.AddWithValue("@QCUser", qcDownload.QCUser);

                    //Calling sp to get production error completed data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        //Writing column headings
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            ws.Cells[0, col].PutValue(sqlReader.GetName(col));
                        }

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                ws.Cells[row, col].PutValue(sqlReader.GetValue(col));
                            }
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

        #region View Existing QC Uploads
        #region View Existing QC Uploads of Project By User
        [HttpGet]
        [Route("ReadExistingProjectUploadsByUser/{CustomerCode}/{ProjectCode}/{QCUser}/{BatchNo?}")]
        public IHttpActionResult ReadExistingProjectUploadsByUser(string CustomerCode, string ProjectCode, string QCUser, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Existing QC Uploads
                List<QCExistingUpload> qcExistingUploadList = new List<QCExistingUpload>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCGetExistingUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@UserID", QCUser);

                    //Calling sp to get list of existing QC Uploads
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCExistingUpload qcExistingUpload = new QCExistingUpload();

                        qcExistingUpload.QCUploadID = Convert.ToInt64(sqlReader["QCUploadID"]);
                        qcExistingUpload.UploadedOn = Convert.ToDateTime(sqlReader["UploadedOn"]);
                        qcExistingUpload.UploadedByUserName = sqlReader["UploadedByUserName"].ToString();
                        qcExistingUpload.Activities = sqlReader["Activities"].ToString();
                        qcExistingUpload.NoOfSKUs = Convert.ToInt32(sqlReader["NoOfSKUs"]);
                        qcExistingUpload.UploadedFileName = sqlReader["UploadedFileName"].ToString();
                        qcExistingUpload.IsQCCompletedCountDownloaded = Convert.ToInt32(sqlReader["IsQCCompletedCountDownloaded"]);

                        qcExistingUploadList.Add(qcExistingUpload);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(qcExistingUploadList);
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

        #region Delete QC Upload
        [HttpPatch]
        [Route("DeleteQCUpload/{id}/{UserID}")]
        public HttpResponseMessage DeleteQCUpload(long id, string UserID)
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
                    cmd.CommandText = "spQCDeleteUpload";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCUploadID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete QC Upload
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        UploadedFileName = arrResult[1];

                        #region Delete Uploaded File
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCUpload/");
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

        #region Download Uploaded file by QC Upload ID
        [HttpGet]
        [Route("downloadfile/{id}")]
        public HttpResponseMessage DownloadUploadedFile(long id)
        {
            try
            {
                string UploadedFileName = string.Empty, fileFullName = string.Empty;
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCUpload/");

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCGetUploadedFileName";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCUploadID", id);

                    //Calling sp to get Allocated FileName
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        UploadedFileName = sqlReader["UploadedFileName"].ToString();
                        fileFullName = Path.Combine(fileUploadedPath, UploadedFileName);

                        //Check whether File exists.
                        if (!File.Exists(fileFullName))
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded File not found");

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

        #region QC Items List
        #region Fetch Moved To QC Customer Codes
        [HttpGet]
        [Route("ReadMovedToQCCustomerCodes")]
        public HttpResponseMessage ReadMovedToQCCustomerCodes()
        {
            try
            {
                DataTable dtDetails = new DataTable();
                List<Customer> lstCustomerCodes = new List<Customer>();
                dtDetails = _BLLQC.ReadMovedToQCCustomerCodes();

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        Customer customer = new Customer();
                        customer.CustomerCode = dtDetails.Rows[i]["CustomerCode"] != DBNull.Value ? dtDetails.Rows[i]["CustomerCode"].ToString().Trim() : "";
                        lstCustomerCodes.Add(customer);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("QCCustomerCodes", new JArray(from p in lstCustomerCodes
                                                                select new JObject(
                                                                new JProperty("CustomerCode", p.CustomerCode)
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

        #region Fetch Moved To QC Project Codes of Customer
        [HttpGet]
        [Route("ReadMovedToQCProjectCodes")]
        public HttpResponseMessage ReadMovedToQCProjectCodes(string CustomerCode)
        {
            try
            {
                DataTable dtDetails = new DataTable();
                List<CustomerCodeProjectCodeBatchNo> lstProjectCodes = new List<CustomerCodeProjectCodeBatchNo>();
                dtDetails = _BLLQC.ReadMovedToQCProjectCodes(CustomerCode);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        CustomerCodeProjectCodeBatchNo customerCodeProjectCodeBatchNo = new CustomerCodeProjectCodeBatchNo();
                        customerCodeProjectCodeBatchNo.ProjectCode = dtDetails.Rows[i]["ProjectCode"] != DBNull.Value ? dtDetails.Rows[i]["ProjectCode"].ToString().Trim() : "";
                        lstProjectCodes.Add(customerCodeProjectCodeBatchNo);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("QCProjectCodes", new JArray(from p in lstProjectCodes
                                                               select new JObject(
                                                                new JProperty("ProjectCode", p.ProjectCode)
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

        #region Fetch Moved To QC Batch Nos. of Project
        [HttpGet]
        [Route("ReadMovedToQCBatchNos")]
        public HttpResponseMessage ReadMovedToQCBatchNos(string CustomerCode, string ProjectCode)
        {
            try
            {
                DataTable dtDetails = new DataTable();
                List<CustomerCodeProjectCodeBatchNo> lstBatchNos = new List<CustomerCodeProjectCodeBatchNo>();
                dtDetails = _BLLQC.ReadMovedToQCBatchNos(CustomerCode, ProjectCode);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        CustomerCodeProjectCodeBatchNo customerCodeProjectCodeBatchNo = new CustomerCodeProjectCodeBatchNo();
                        customerCodeProjectCodeBatchNo.BatchNo = dtDetails.Rows[i]["BatchNo"] != DBNull.Value ? dtDetails.Rows[i]["BatchNo"].ToString().Trim() : "";
                        lstBatchNos.Add(customerCodeProjectCodeBatchNo);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("QCBatchNos", new JArray(from b in lstBatchNos
                                                           select new JObject(
                                                                new JProperty("BatchNo", b.BatchNo)
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

        #region Fetch Moved To QC Noun/Modifiers
        [HttpGet]
        [Route("ReadMovedToQCNounModifiers")]
        public HttpResponseMessage ReadMovedToQCNounModifiers(string CustomerCode, string ProjectCode, string BatchNo="")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                DataTable dtDetails = new DataTable();
                List<MovedToQCNounModifier> lstNounModifiers = new List<MovedToQCNounModifier>();
                dtDetails = _BLLQC.ReadMovedToQCNounModifiers(CustomerCode, ProjectCode, BatchNo);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        MovedToQCNounModifier movedToQCNounModifier = new MovedToQCNounModifier();
                        movedToQCNounModifier.Noun = dtDetails.Rows[i]["Noun"] != DBNull.Value ? dtDetails.Rows[i]["Noun"].ToString().Trim() : "";
                        movedToQCNounModifier.Modifier = dtDetails.Rows[i]["Modifier"] != DBNull.Value ? dtDetails.Rows[i]["Modifier"].ToString().Trim() : "";
                        movedToQCNounModifier.CountOfNMItems = dtDetails.Rows[i]["CountOfNMItems"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["CountOfNMItems"]) : 0;
                        movedToQCNounModifier.TotalCountOfItems = dtDetails.Rows[i]["TotalCountOfItems"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["TotalCountOfItems"]) : 0;
                        movedToQCNounModifier.ProjectScope = dtDetails.Rows[i]["ProjectScope"] != DBNull.Value ? dtDetails.Rows[i]["ProjectScope"].ToString().Trim() : "";
                        lstNounModifiers.Add(movedToQCNounModifier);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("ProjectScope", dtDetails.Rows[0]["ProjectScope"].ToString()),
                    new JProperty("TotalCountOfItems", Convert.ToInt32(dtDetails.Rows[0]["TotalCountOfItems"])),
                    new JProperty("NounModifiers", new JArray(from nm in lstNounModifiers
                                                              select new JObject(
                                                                new JProperty("NounModifier", nm.Noun + " / " + nm.Modifier),
                                                                new JProperty("CountOfNMItems", nm.CountOfNMItems)
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

        //#region Fetch Moved To QC Project Items
        //[HttpGet]
        //[Route("ReadMovedToQCProjectItems")]
        //public HttpResponseMessage ReadMovedToQCProjectItems(string CustomerCode, string ProjectCode, string BatchNo = "", string Noun = "", string Modifier = "")
        //{
        //    try
        //    {
        //        //Swagger fix
        //        if (BatchNo == "{BatchNo}")
        //            BatchNo = "";

        //        DataTable dtDetails = new DataTable();
        //        List<QCItem> itemsList = new List<QCItem>();
        //        dtDetails = _BLLQC.ReadMovedToQCProjectItems(CustomerCode, ProjectCode, BatchNo, Noun, Modifier);

        //        if (dtDetails.Rows.Count > 0)
        //        {
        //            for (int i = 0; i < dtDetails.Rows.Count; i++)
        //            {
        //                QCItem item = new QCItem();

        //                if (dtDetails.Rows[i]["QCItemID"] != DBNull.Value)
        //                    item.QCItemID = Convert.ToInt64(dtDetails.Rows[i]["QCItemID"]);
        //                item.ProductionItemID = Convert.ToInt64(dtDetails.Rows[i]["ProductionItemID"]);
        //                if (dtDetails.Rows[i]["QCTestNo"] != DBNull.Value)
        //                    item.QCTestNo = Convert.ToInt32(dtDetails.Rows[i]["QCTestNo"]);
        //                item.UniqueID = dtDetails.Rows[i]["UniqueID"].ToString().Trim().ToString();
        //                item.ShortDescription = dtDetails.Rows[i]["ShortDescription"] != DBNull.Value ? dtDetails.Rows[i]["ShortDescription"].ToString().Trim() : "";
        //                item.LongDescription = dtDetails.Rows[i]["LongDescription"] != DBNull.Value ? dtDetails.Rows[i]["LongDescription"].ToString().Trim() : "";
        //                item.UOM = dtDetails.Rows[i]["UOM"] != DBNull.Value ? dtDetails.Rows[i]["UOM"].ToString().Trim() : "";
        //                item.ProductionUser = dtDetails.Rows[i]["ProductionUser"] != DBNull.Value ? dtDetails.Rows[i]["ProductionUser"].ToString().Trim() : "";
        //                item.MFRName = dtDetails.Rows[i]["MFRName"] != DBNull.Value ? dtDetails.Rows[i]["MFRName"].ToString().Trim() : "";
        //                item.MFRPN = dtDetails.Rows[i]["MFRPN"] != DBNull.Value ? dtDetails.Rows[i]["MFRPN"].ToString().Trim() : "";
        //                item.VendorName = dtDetails.Rows[i]["VendorName"] != DBNull.Value ? dtDetails.Rows[i]["VendorName"].ToString().Trim() : "";
        //                item.VendorPN = dtDetails.Rows[i]["VendorPN"] != DBNull.Value ? dtDetails.Rows[i]["VendorPN"].ToString().Trim() : "";
        //                item.NewShortDescription = dtDetails.Rows[i]["NewShortDescription"] != DBNull.Value ? dtDetails.Rows[i]["NewShortDescription"].ToString().Trim() : "";
        //                item.NewLongDescription = dtDetails.Rows[i]["NewLongDescription"] != DBNull.Value ? dtDetails.Rows[i]["NewLongDescription"].ToString().Trim() : "";
        //                item.MissingWords = dtDetails.Rows[i]["MissingWords"] != DBNull.Value ? dtDetails.Rows[i]["MissingWords"].ToString().Trim() : "";
        //                item.QCUser = dtDetails.Rows[i]["QCUser"] != DBNull.Value ? dtDetails.Rows[i]["QCUser"].ToString().Trim() : "";
        //                item.QCStatus = dtDetails.Rows[i]["QCStatus"] != DBNull.Value ? dtDetails.Rows[i]["QCStatus"].ToString().Trim() : "";
        //                item.Noun = dtDetails.Rows[i]["Noun"] != DBNull.Value ? dtDetails.Rows[i]["Noun"].ToString().Trim() : "";
        //                item.Modifier = dtDetails.Rows[i]["Modifier"] != DBNull.Value ? dtDetails.Rows[i]["Modifier"].ToString().Trim() : "";
        //                item.QCLevel = dtDetails.Rows[i]["Level"] != DBNull.Value ? dtDetails.Rows[i]["Level"].ToString().Trim() : "";
        //                item.IsQCEditable = dtDetails.Rows[i]["IsQCEditable"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsQCEditable"]) : false;
        //                itemsList.Add(item);
        //            }

        //            JObject PEReport = new JObject(new JProperty("Success", 1),
        //            new JProperty("RecordCount", dtDetails.Rows.Count),
        //            new JProperty("ItemsList", new JArray(from i in itemsList
        //                                                  select new JObject(
        //                                                        new JProperty("QCItemID", i.QCItemID),
        //                                                        new JProperty("ProductionItemID", i.ProductionItemID),
        //                                                        new JProperty("QCTestNo", i.QCTestNo),
        //                                                        new JProperty("UniqueID", i.UniqueID),
        //                                                        new JProperty("ShortDescription", i.ShortDescription),
        //                                                        new JProperty("LongDescription", i.LongDescription),
        //                                                        new JProperty("UOM", i.UOM),
        //                                                        new JProperty("ProductionUser", i.ProductionUser),
        //                                                        new JProperty("MFRName", i.MFRName),
        //                                                        new JProperty("MFRPN", i.MFRPN),
        //                                                        new JProperty("VendorName", i.VendorName),
        //                                                        new JProperty("VendorPN", i.VendorPN),
        //                                                        new JProperty("NewShortDescription", i.NewShortDescription),
        //                                                        new JProperty("NewLongDescription", i.NewLongDescription),
        //                                                        new JProperty("MissingWords", i.MissingWords),
        //                                                        new JProperty("QCUser", i.QCUser),
        //                                                        new JProperty("QCStatus", i.QCStatus),
        //                                                        new JProperty("Noun", i.Noun),
        //                                                        new JProperty("Modifier", i.Modifier),
        //                                                        new JProperty("Level", i.QCLevel),
        //                                                        new JProperty("IsQCEditable", i.IsQCEditable)
        //                                                        ))));
        //            return Request.CreateResponse(HttpStatusCode.OK, PEReport);
        //        }
        //        else
        //        {
        //            Result objResult = new Result();
        //            objResult.Msg = "No data found";
        //            objResult.Success = 0;
        //            return Request.CreateResponse(HttpStatusCode.OK, objResult);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // Log the exception
        //        ExceptionLogging.SendExceptionToDB(ex);
        //        Result objResult = new Result();
        //        objResult.Success = 0;
        //        objResult.Msg = ex.Message;
        //        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
        //    }
        //}
        //#endregion

        #region Fetch Moved To QC Project Items (without using BLL/DAL)
        [HttpGet]
        [Route("ReadMovedToQCProjectItems")]
        public IHttpActionResult ReadMovedToQCProjectItems(string CustomerCode, string ProjectCode, string BatchNo = "", string Noun = "", string Modifier = "", int PageNo = 1, int PageSize = 100)
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                List<QCItem> itemsList = new List<QCItem>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMovedToQCProjectItems";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    cmd.Parameters.AddWithValue("@PageNo", PageNo);
                    cmd.Parameters.AddWithValue("@PageSize", PageSize);
                    cmd.Parameters.AddWithValue("@IsToFetchALLDetails", false);
                    cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;

                    cmd.CommandTimeout = 500;
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCItem item = new QCItem();
                        item.QCItemID = sqlReader["QCItemID"] == DBNull.Value ? null : (Int64?)Convert.ToInt64(sqlReader["QCItemID"]);
                        item.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        item.QCTestNo = sqlReader["QCTestNo"] == DBNull.Value ? null : (Int32?)Convert.ToInt32(sqlReader["QCTestNo"]);
                        item.UniqueID = sqlReader["UniqueID"].ToString();
                        item.ShortDescription = sqlReader["ShortDescription"].ToString();
                        item.LongDescription = sqlReader["LongDescription"].ToString();
                        item.UOM = sqlReader["UOM"].ToString();
                        item.MFRName = sqlReader["MFRName"].ToString();
                        item.MFRPN = sqlReader["MFRPN"].ToString();
                        item.VendorName = sqlReader["VendorName"].ToString();
                        item.VendorPN = sqlReader["VendorPN"].ToString();
                        item.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        item.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        item.MissingWords = sqlReader["MissingWords"].ToString();
                        item.Noun = sqlReader["Noun"].ToString();
                        item.Modifier = sqlReader["Modifier"].ToString();
                        item.QCUser = sqlReader["QCUser"].ToString();
                        item.ProductionUser = sqlReader["ProductionUser"].ToString();
                        item.QCStatus = sqlReader["QCStatus"].ToString();
                        item.QCLevel = sqlReader["Level"].ToString();
                        item.IsQCEditable = Convert.ToBoolean(sqlReader["IsQCEditable"]);
                        item.TotalRowsCount = Convert.ToInt32(sqlReader["TotalCount"]);
                        itemsList.Add(item);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(itemsList);

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

        #region Export the List of Moved To QC Noun / Modifiers to Excel
        [HttpGet]
        [Route("ExportMovedToQCNounModifiersListToExcel")]
        public HttpResponseMessage ExportMovedToQCNounModifiersListToExcel(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string FileName = "Moved To QC Noun-Modifiers List.xlsx";

                DataTable dtDetails = new DataTable();
                List<MovedToQCNounModifier> lstNounModifiers = new List<MovedToQCNounModifier>();
                dtDetails = _BLLQC.ReadMovedToQCNounModifiers(CustomerCode, ProjectCode, BatchNo);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        MovedToQCNounModifier movedToQCNounModifier = new MovedToQCNounModifier();
                        movedToQCNounModifier.Noun = dtDetails.Rows[i]["Noun"] != DBNull.Value ? dtDetails.Rows[i]["Noun"].ToString().Trim() : "";
                        movedToQCNounModifier.Modifier = dtDetails.Rows[i]["Modifier"] != DBNull.Value ? dtDetails.Rows[i]["Modifier"].ToString().Trim() : "";
                        movedToQCNounModifier.CountOfNMItems = dtDetails.Rows[i]["CountOfNMItems"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["CountOfNMItems"]) : 0;
                        lstNounModifiers.Add(movedToQCNounModifier);
                    }

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
                    ws.Cells[0, 1].PutValue("Noun");
                    ws.Cells[0, 2].PutValue("Modifier");
                    ws.Cells[0, 3].PutValue("Count of Items");
 
                    for (int c = 0; c <= 3; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (MovedToQCNounModifier nm in lstNounModifiers)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(row);
                        ws.Cells[row, 1].PutValue(nm.Noun);
                        ws.Cells[row, 2].PutValue(nm.Modifier);
                        ws.Cells[row, 3].PutValue(nm.CountOfNMItems);
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
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

        #region Export the List of Moved To QC Items List to Excel
        [HttpGet]
        [Route("ExportMovedToQCItemsListToExcel")]
        public HttpResponseMessage ExportMovedToQCItemsListToExcel(string CustomerCode, string ProjectCode, string BatchNo = "", string Noun = "", string Modifier = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string FileName = "Moved To QC Items List.xlsx";

                DataTable dtDetails = new DataTable();
                List<ProductionItem> itemsList = new List<ProductionItem>();
                dtDetails = _BLLQC.ReadMovedToQCProjectItems(CustomerCode, ProjectCode, BatchNo, Noun, Modifier,1,100000000,true);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        ProductionItem item = new ProductionItem();

                        item.ProductionItemID = Convert.ToInt64(dtDetails.Rows[i]["ProductionItemID"]);
                        item.UniqueID = dtDetails.Rows[i]["UniqueID"].ToString().Trim().ToString();
                        item.ShortDescription = dtDetails.Rows[i]["ShortDescription"] != DBNull.Value ? dtDetails.Rows[i]["ShortDescription"].ToString().Trim() : "";
                        item.LongDescription = dtDetails.Rows[i]["LongDescription"] != DBNull.Value ? dtDetails.Rows[i]["LongDescription"].ToString().Trim() : "";
                        item.UOM = dtDetails.Rows[i]["UOM"] != DBNull.Value ? dtDetails.Rows[i]["UOM"].ToString().Trim() : "";
                        item.ProductionUser = dtDetails.Rows[i]["ProductionUser"] != DBNull.Value ? dtDetails.Rows[i]["ProductionUser"].ToString().Trim() : "";
                        item.QCUser = dtDetails.Rows[i]["QCUser"] != DBNull.Value ? dtDetails.Rows[i]["QCUser"].ToString().Trim() : "";
                        item.QCStatus = dtDetails.Rows[i]["QCStatus"] != DBNull.Value ? dtDetails.Rows[i]["QCStatus"].ToString().Trim() : "";
                        item.QCTestNo = dtDetails.Rows[i]["QCTestNo"] != DBNull.Value ? dtDetails.Rows[i]["QCTestNo"].ToString().Trim() : "";
                        item.MFRName = dtDetails.Rows[i]["MFRName"] != DBNull.Value ? dtDetails.Rows[i]["MFRName"].ToString().Trim() : "";
                        item.MFRPN = dtDetails.Rows[i]["MFRPN"] != DBNull.Value ? dtDetails.Rows[i]["MFRPN"].ToString().Trim() : "";
                        item.VendorName = dtDetails.Rows[i]["VendorName"] != DBNull.Value ? dtDetails.Rows[i]["VendorName"].ToString().Trim() : "";
                        item.VendorPN = dtDetails.Rows[i]["VendorPN"] != DBNull.Value ? dtDetails.Rows[i]["VendorPN"].ToString().Trim() : "";
                        item.NewShortDescription = dtDetails.Rows[i]["NewShortDescription"] != DBNull.Value ? dtDetails.Rows[i]["NewShortDescription"].ToString().Trim() : "";
                        item.NewLongDescription = dtDetails.Rows[i]["NewLongDescription"] != DBNull.Value ? dtDetails.Rows[i]["NewLongDescription"].ToString().Trim() : "";
                        item.MissingWords = dtDetails.Rows[i]["MissingWords"] != DBNull.Value ? dtDetails.Rows[i]["MissingWords"].ToString().Trim() : "";
                        item.Noun = dtDetails.Rows[i]["Noun"] != DBNull.Value ? dtDetails.Rows[i]["Noun"].ToString().Trim() : "";
                        item.Modifier = dtDetails.Rows[i]["Modifier"] != DBNull.Value ? dtDetails.Rows[i]["Modifier"].ToString().Trim() : "";
                        item.Level = dtDetails.Rows[i]["Level"] != DBNull.Value ? dtDetails.Rows[i]["Level"].ToString().Trim() : "";
                        item.MFRName1 = dtDetails.Rows[i]["MFRName1"] != DBNull.Value ? dtDetails.Rows[i]["MFRName1"].ToString().Trim() : "";
                        item.MFRPN1 = dtDetails.Rows[i]["MFRPN1"] != DBNull.Value ? dtDetails.Rows[i]["MFRPN1"].ToString().Trim() : "";
                        item.MFRName2 = dtDetails.Rows[i]["MFRName2"] != DBNull.Value ? dtDetails.Rows[i]["MFRName2"].ToString().Trim() : "";
                        item.MFRPN2 = dtDetails.Rows[i]["MFRPN2"] != DBNull.Value ? dtDetails.Rows[i]["MFRPN2"].ToString().Trim() : "";
                        item.MFRName3 = dtDetails.Rows[i]["MFRName3"] != DBNull.Value ? dtDetails.Rows[i]["MFRName3"].ToString().Trim() : "";
                        item.MFRPN3 = dtDetails.Rows[i]["MFRPN3"] != DBNull.Value ? dtDetails.Rows[i]["MFRPN3"].ToString().Trim() : "";
                        item.VendorName1 = dtDetails.Rows[i]["VendorName1"] != DBNull.Value ? dtDetails.Rows[i]["VendorName1"].ToString().Trim() : "";
                        item.VendorPN1 = dtDetails.Rows[i]["VendorPN1"] != DBNull.Value ? dtDetails.Rows[i]["VendorPN1"].ToString().Trim() : "";
                        item.VendorName2 = dtDetails.Rows[i]["VendorName2"] != DBNull.Value ? dtDetails.Rows[i]["VendorName2"].ToString().Trim() : "";
                        item.VendorPN2 = dtDetails.Rows[i]["VendorPN2"] != DBNull.Value ? dtDetails.Rows[i]["VendorPN2"].ToString().Trim() : "";
                        item.VendorName3 = dtDetails.Rows[i]["VendorName3"] != DBNull.Value ? dtDetails.Rows[i]["VendorName3"].ToString().Trim() : "";
                        item.VendorPN3 = dtDetails.Rows[i]["VendorPN3"] != DBNull.Value ? dtDetails.Rows[i]["VendorPN3"].ToString().Trim() : "";
                        item.AdditionalInfo = dtDetails.Rows[i]["AdditionalInfo"] != DBNull.Value ? dtDetails.Rows[i]["AdditionalInfo"].ToString().Trim() : "";
                        item.AdditionalInfoFromWeb = dtDetails.Rows[i]["AdditionalInfoFromWeb"] != DBNull.Value ? dtDetails.Rows[i]["AdditionalInfoFromWeb"].ToString().Trim() : "";
                        item.UNSPSCCode = dtDetails.Rows[i]["UNSPSCCode"] != DBNull.Value ? dtDetails.Rows[i]["UNSPSCCode"].ToString().Trim() : "";
                        item.UNSPSCCategory = dtDetails.Rows[i]["UNSPSCCategory"] != DBNull.Value ? dtDetails.Rows[i]["UNSPSCCategory"].ToString().Trim() : "";
                        item.WebRefURL1 = dtDetails.Rows[i]["WebRefURL1"] != DBNull.Value ? dtDetails.Rows[i]["WebRefURL1"].ToString().Trim() : "";
                        item.WebRefURL2 = dtDetails.Rows[i]["WebRefURL2"] != DBNull.Value ? dtDetails.Rows[i]["WebRefURL2"].ToString().Trim() : "";
                        item.WebRefURL3 = dtDetails.Rows[i]["WebRefURL3"] != DBNull.Value ? dtDetails.Rows[i]["WebRefURL3"].ToString().Trim() : "";
                        item.PDFURL = dtDetails.Rows[i]["PDFURL"] != DBNull.Value ? dtDetails.Rows[i]["PDFURL"].ToString().Trim() : "";
                        item.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        item.Query = dtDetails.Rows[i]["Query"] != DBNull.Value ? dtDetails.Rows[i]["Query"].ToString().Trim() : "";
                        item.Application = dtDetails.Rows[i]["Application"] != DBNull.Value ? dtDetails.Rows[i]["Application"].ToString().Trim() : "";
                        item.DWG = dtDetails.Rows[i]["DWG"] != DBNull.Value ? dtDetails.Rows[i]["DWG"].ToString().Trim() : "";
                        item.POS = dtDetails.Rows[i]["POS"] != DBNull.Value ? dtDetails.Rows[i]["POS"].ToString().Trim() : "";
                        item.ItemNo = dtDetails.Rows[i]["ItemNo"] != DBNull.Value ? dtDetails.Rows[i]["ItemNo"].ToString().Trim() : "";
                        item.SerialNo = dtDetails.Rows[i]["SerialNo"] != DBNull.Value ? dtDetails.Rows[i]["SerialNo"].ToString().Trim() : "";
                        item.OtherNo = dtDetails.Rows[i]["OtherNo"] != DBNull.Value ? dtDetails.Rows[i]["OtherNo"].ToString().Trim() : "";
                        item.KKSCode = dtDetails.Rows[i]["KKSCode"] != DBNull.Value ? dtDetails.Rows[i]["KKSCode"].ToString().Trim() : "";
                        item.AssemblyOrPart = dtDetails.Rows[i]["AssemblyOrPart"] != DBNull.Value ? dtDetails.Rows[i]["AssemblyOrPart"].ToString().Trim() : "";
                        item.BOM = dtDetails.Rows[i]["BOM"] != DBNull.Value ? dtDetails.Rows[i]["BOM"].ToString().Trim() : "";
                        item.GreenItems = dtDetails.Rows[i]["GreenItems"] != DBNull.Value ? dtDetails.Rows[i]["GreenItems"].ToString().Trim() : "";
                        item.AttributeName1 = dtDetails.Rows[i]["AttributeName1"].ToString();
                        item.AttributeValue1 = dtDetails.Rows[i]["AttributeValue1"].ToString();
                        item.AttributeName2 = dtDetails.Rows[i]["AttributeName2"].ToString();
                        item.AttributeValue2 = dtDetails.Rows[i]["AttributeValue2"].ToString();
                        item.AttributeName3 = dtDetails.Rows[i]["AttributeName3"].ToString();
                        item.AttributeValue3 = dtDetails.Rows[i]["AttributeValue3"].ToString();
                        item.AttributeName4 = dtDetails.Rows[i]["AttributeName4"].ToString();
                        item.AttributeValue4 = dtDetails.Rows[i]["AttributeValue4"].ToString();
                        item.AttributeName5 = dtDetails.Rows[i]["AttributeName5"].ToString();
                        item.AttributeValue5 = dtDetails.Rows[i]["AttributeValue5"].ToString();
                        item.AttributeName6 = dtDetails.Rows[i]["AttributeName6"].ToString();
                        item.AttributeValue6 = dtDetails.Rows[i]["AttributeValue6"].ToString();
                        item.AttributeName7 = dtDetails.Rows[i]["AttributeName7"].ToString();
                        item.AttributeValue7 = dtDetails.Rows[i]["AttributeValue7"].ToString();
                        item.AttributeName8 = dtDetails.Rows[i]["AttributeName8"].ToString();
                        item.AttributeValue8 = dtDetails.Rows[i]["AttributeValue8"].ToString();
                        item.AttributeName9 = dtDetails.Rows[i]["AttributeName9"].ToString();
                        item.AttributeValue9 = dtDetails.Rows[i]["AttributeValue9"].ToString();
                        item.AttributeName10 = dtDetails.Rows[i]["AttributeName10"].ToString();
                        item.AttributeValue10 = dtDetails.Rows[i]["AttributeValue10"].ToString();
                        item.AttributeName11 = dtDetails.Rows[i]["AttributeName11"].ToString();
                        item.AttributeValue11 = dtDetails.Rows[i]["AttributeValue11"].ToString();
                        item.AttributeName12 = dtDetails.Rows[i]["AttributeName12"].ToString();
                        item.AttributeValue12 = dtDetails.Rows[i]["AttributeValue12"].ToString();
                        item.AttributeName13 = dtDetails.Rows[i]["AttributeName13"].ToString();
                        item.AttributeValue13 = dtDetails.Rows[i]["AttributeValue13"].ToString();
                        item.AttributeName14 = dtDetails.Rows[i]["AttributeName14"].ToString();
                        item.AttributeValue14 = dtDetails.Rows[i]["AttributeValue14"].ToString();
                        item.AttributeName15 = dtDetails.Rows[i]["AttributeName15"].ToString();
                        item.AttributeValue15 = dtDetails.Rows[i]["AttributeValue15"].ToString();
                        item.AttributeName16 = dtDetails.Rows[i]["AttributeName16"].ToString();
                        item.AttributeValue16 = dtDetails.Rows[i]["AttributeValue16"].ToString();
                        item.AttributeName17 = dtDetails.Rows[i]["AttributeName17"].ToString();
                        item.AttributeValue17 = dtDetails.Rows[i]["AttributeValue17"].ToString();
                        item.AttributeName18 = dtDetails.Rows[i]["AttributeName18"].ToString();
                        item.AttributeValue18 = dtDetails.Rows[i]["AttributeValue18"].ToString();
                        item.AttributeName19 = dtDetails.Rows[i]["AttributeName19"].ToString();
                        item.AttributeValue19 = dtDetails.Rows[i]["AttributeValue19"].ToString();
                        item.AttributeName20 = dtDetails.Rows[i]["AttributeName20"].ToString();
                        item.AttributeValue20 = dtDetails.Rows[i]["AttributeValue20"].ToString();
                        item.AttributeName21 = dtDetails.Rows[i]["AttributeName21"].ToString();
                        item.AttributeValue21 = dtDetails.Rows[i]["AttributeValue21"].ToString();
                        item.AttributeName22 = dtDetails.Rows[i]["AttributeName22"].ToString();
                        item.AttributeValue22 = dtDetails.Rows[i]["AttributeValue22"].ToString();
                        item.AttributeName23 = dtDetails.Rows[i]["AttributeName23"].ToString();
                        item.AttributeValue23 = dtDetails.Rows[i]["AttributeValue23"].ToString();
                        item.AttributeName24 = dtDetails.Rows[i]["AttributeName24"].ToString();
                        item.AttributeValue24 = dtDetails.Rows[i]["AttributeValue24"].ToString();
                        item.AttributeName25 = dtDetails.Rows[i]["AttributeName25"].ToString();
                        item.AttributeValue25 = dtDetails.Rows[i]["AttributeValue25"].ToString();
                        item.AttributeName26 = dtDetails.Rows[i]["AttributeName26"].ToString();
                        item.AttributeValue26 = dtDetails.Rows[i]["AttributeValue26"].ToString();
                        item.AttributeName27 = dtDetails.Rows[i]["AttributeName27"].ToString();
                        item.AttributeValue27 = dtDetails.Rows[i]["AttributeValue27"].ToString();
                        item.AttributeName28 = dtDetails.Rows[i]["AttributeName28"].ToString();
                        item.AttributeValue28 = dtDetails.Rows[i]["AttributeValue28"].ToString();
                        item.AttributeName29 = dtDetails.Rows[i]["AttributeName29"].ToString();
                        item.AttributeValue29 = dtDetails.Rows[i]["AttributeValue29"].ToString();
                        item.AttributeName30 = dtDetails.Rows[i]["AttributeName30"].ToString();
                        item.AttributeValue30 = dtDetails.Rows[i]["AttributeValue30"].ToString();
                        item.AttributeName31 = dtDetails.Rows[i]["AttributeName31"].ToString();
                        item.AttributeValue31 = dtDetails.Rows[i]["AttributeValue31"].ToString();
                        item.AttributeName32 = dtDetails.Rows[i]["AttributeName32"].ToString();
                        item.AttributeValue32 = dtDetails.Rows[i]["AttributeValue32"].ToString();
                        item.AttributeName33 = dtDetails.Rows[i]["AttributeName33"].ToString();
                        item.AttributeValue33 = dtDetails.Rows[i]["AttributeValue33"].ToString();
                        item.AttributeName34 = dtDetails.Rows[i]["AttributeName34"].ToString();
                        item.AttributeValue34 = dtDetails.Rows[i]["AttributeValue34"].ToString();
                        item.AttributeName35 = dtDetails.Rows[i]["AttributeName35"].ToString();
                        item.AttributeValue35 = dtDetails.Rows[i]["AttributeValue35"].ToString();
                        item.AttributeName36 = dtDetails.Rows[i]["AttributeName36"].ToString();
                        item.AttributeValue36 = dtDetails.Rows[i]["AttributeValue36"].ToString();
                        item.AttributeName37 = dtDetails.Rows[i]["AttributeName37"].ToString();
                        item.AttributeValue37 = dtDetails.Rows[i]["AttributeValue37"].ToString();
                        item.AttributeName38 = dtDetails.Rows[i]["AttributeName38"].ToString();
                        item.AttributeValue38 = dtDetails.Rows[i]["AttributeValue38"].ToString();
                        item.AttributeName39 = dtDetails.Rows[i]["AttributeName39"].ToString();
                        item.AttributeValue39 = dtDetails.Rows[i]["AttributeValue39"].ToString();
                        item.AttributeName40 = dtDetails.Rows[i]["AttributeName40"].ToString();
                        item.AttributeValue40 = dtDetails.Rows[i]["AttributeValue40"].ToString();
                        item.AttributeName41 = dtDetails.Rows[i]["AttributeName41"].ToString();
                        item.AttributeValue41 = dtDetails.Rows[i]["AttributeValue41"].ToString();
                        item.AttributeName42 = dtDetails.Rows[i]["AttributeName42"].ToString();
                        item.AttributeValue42 = dtDetails.Rows[i]["AttributeValue42"].ToString();
                        item.AttributeName43 = dtDetails.Rows[i]["AttributeName43"].ToString();
                        item.AttributeValue43 = dtDetails.Rows[i]["AttributeValue43"].ToString();
                        item.AttributeName44 = dtDetails.Rows[i]["AttributeName44"].ToString();
                        item.AttributeValue44 = dtDetails.Rows[i]["AttributeValue44"].ToString();
                        item.AttributeName45 = dtDetails.Rows[i]["AttributeName45"].ToString();
                        item.AttributeValue45 = dtDetails.Rows[i]["AttributeValue45"].ToString();
                        item.AttributeName46 = dtDetails.Rows[i]["AttributeName46"].ToString();
                        item.AttributeValue46 = dtDetails.Rows[i]["AttributeValue46"].ToString();
                        item.AttributeName47 = dtDetails.Rows[i]["AttributeName47"].ToString();
                        item.AttributeValue47 = dtDetails.Rows[i]["AttributeValue47"].ToString();
                        item.AttributeName48 = dtDetails.Rows[i]["AttributeName48"].ToString();
                        item.AttributeValue48 = dtDetails.Rows[i]["AttributeValue48"].ToString();
                        item.AttributeName49 = dtDetails.Rows[i]["AttributeName49"].ToString();
                        item.AttributeValue49 = dtDetails.Rows[i]["AttributeValue49"].ToString();
                        item.AttributeName50 = dtDetails.Rows[i]["AttributeName50"].ToString();
                        item.AttributeValue50 = dtDetails.Rows[i]["AttributeValue50"].ToString();
                        item.CustomColumnName1 = dtDetails.Rows[i]["CustomColumnName1"].ToString();
                        item.CustomColumnName1Value = dtDetails.Rows[i]["CustomColumnName1Value"].ToString();
                        item.CustomColumnName2 = dtDetails.Rows[i]["CustomColumnName2"].ToString();
                        item.CustomColumnName2Value = dtDetails.Rows[i]["CustomColumnName2Value"].ToString();
                        item.CustomColumnName3 = dtDetails.Rows[i]["CustomColumnName3"].ToString();
                        item.CustomColumnName3Value = dtDetails.Rows[i]["CustomColumnName3Value"].ToString();

                        itemsList.Add(item);
                    }

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
                    int colNo = 51;
                    for (int attributeNo = 1; attributeNo <= 50; attributeNo++)
                    {
                        ws.Cells[0, colNo].PutValue("Attribute Name" + attributeNo.ToString());
                        ws.Cells[0, colNo + 1].PutValue("Attribute Value" + attributeNo.ToString());
                        colNo += 2;
                    }
                    if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName1))
                        ws.Cells[0, 151].PutValue(itemsList.FirstOrDefault().CustomColumnName1);
                    if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName2))
                        ws.Cells[0, 152].PutValue(itemsList.FirstOrDefault().CustomColumnName2);
                    if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName3))
                        ws.Cells[0, 153].PutValue(itemsList.FirstOrDefault().CustomColumnName3);
                    
                    for (int c = 0; c <= ws.Cells.MaxColumn; c++)
                           ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (ProductionItem pi in itemsList)
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
                        ws.Cells[row, 51].PutValue(pi.AttributeName1);
                        ws.Cells[row, 52].PutValue(pi.AttributeValue1);
                        ws.Cells[row, 53].PutValue(pi.AttributeName2);
                        ws.Cells[row, 54].PutValue(pi.AttributeValue2);
                        ws.Cells[row, 55].PutValue(pi.AttributeName3);
                        ws.Cells[row, 56].PutValue(pi.AttributeValue3);
                        ws.Cells[row, 57].PutValue(pi.AttributeName4);
                        ws.Cells[row, 58].PutValue(pi.AttributeValue4);
                        ws.Cells[row, 59].PutValue(pi.AttributeName5);
                        ws.Cells[row, 60].PutValue(pi.AttributeValue5);
                        ws.Cells[row, 61].PutValue(pi.AttributeName6);
                        ws.Cells[row, 62].PutValue(pi.AttributeValue6);
                        ws.Cells[row, 63].PutValue(pi.AttributeName7);
                        ws.Cells[row, 64].PutValue(pi.AttributeValue7);
                        ws.Cells[row, 65].PutValue(pi.AttributeName8);
                        ws.Cells[row, 66].PutValue(pi.AttributeValue8);
                        ws.Cells[row, 67].PutValue(pi.AttributeName9);
                        ws.Cells[row, 68].PutValue(pi.AttributeValue9);
                        ws.Cells[row, 69].PutValue(pi.AttributeName10);
                        ws.Cells[row, 70].PutValue(pi.AttributeValue10);
                        ws.Cells[row, 71].PutValue(pi.AttributeName11);
                        ws.Cells[row, 72].PutValue(pi.AttributeValue11);
                        ws.Cells[row, 73].PutValue(pi.AttributeName12);
                        ws.Cells[row, 74].PutValue(pi.AttributeValue12);
                        ws.Cells[row, 75].PutValue(pi.AttributeName13);
                        ws.Cells[row, 76].PutValue(pi.AttributeValue13);
                        ws.Cells[row, 77].PutValue(pi.AttributeName14);
                        ws.Cells[row, 78].PutValue(pi.AttributeValue14);
                        ws.Cells[row, 79].PutValue(pi.AttributeName15);
                        ws.Cells[row, 80].PutValue(pi.AttributeValue15);
                        ws.Cells[row, 81].PutValue(pi.AttributeName16);
                        ws.Cells[row, 82].PutValue(pi.AttributeValue16);
                        ws.Cells[row, 83].PutValue(pi.AttributeName17);
                        ws.Cells[row, 84].PutValue(pi.AttributeValue17);
                        ws.Cells[row, 85].PutValue(pi.AttributeName18);
                        ws.Cells[row, 86].PutValue(pi.AttributeValue18);
                        ws.Cells[row, 87].PutValue(pi.AttributeName19);
                        ws.Cells[row, 88].PutValue(pi.AttributeValue19);
                        ws.Cells[row, 89].PutValue(pi.AttributeName20);
                        ws.Cells[row, 90].PutValue(pi.AttributeValue20);
                        ws.Cells[row, 91].PutValue(pi.AttributeName21);
                        ws.Cells[row, 92].PutValue(pi.AttributeValue21);
                        ws.Cells[row, 93].PutValue(pi.AttributeName22);
                        ws.Cells[row, 94].PutValue(pi.AttributeValue22);
                        ws.Cells[row, 95].PutValue(pi.AttributeName23);
                        ws.Cells[row, 96].PutValue(pi.AttributeValue23);
                        ws.Cells[row, 97].PutValue(pi.AttributeName24);
                        ws.Cells[row, 98].PutValue(pi.AttributeValue24);
                        ws.Cells[row, 99].PutValue(pi.AttributeName25);
                        ws.Cells[row, 100].PutValue(pi.AttributeValue25);
                        ws.Cells[row, 101].PutValue(pi.AttributeName26);
                        ws.Cells[row, 102].PutValue(pi.AttributeValue26);
                        ws.Cells[row, 103].PutValue(pi.AttributeName27);
                        ws.Cells[row, 104].PutValue(pi.AttributeValue27);
                        ws.Cells[row, 105].PutValue(pi.AttributeName28);
                        ws.Cells[row, 106].PutValue(pi.AttributeValue28);
                        ws.Cells[row, 107].PutValue(pi.AttributeName29);
                        ws.Cells[row, 108].PutValue(pi.AttributeValue29);
                        ws.Cells[row, 109].PutValue(pi.AttributeName30);
                        ws.Cells[row, 110].PutValue(pi.AttributeValue30);
                        ws.Cells[row, 111].PutValue(pi.AttributeName31);
                        ws.Cells[row, 112].PutValue(pi.AttributeValue31);
                        ws.Cells[row, 113].PutValue(pi.AttributeName32);
                        ws.Cells[row, 114].PutValue(pi.AttributeValue32);
                        ws.Cells[row, 115].PutValue(pi.AttributeName33);
                        ws.Cells[row, 116].PutValue(pi.AttributeValue33);
                        ws.Cells[row, 117].PutValue(pi.AttributeName34);
                        ws.Cells[row, 118].PutValue(pi.AttributeValue34);
                        ws.Cells[row, 119].PutValue(pi.AttributeName35);
                        ws.Cells[row, 120].PutValue(pi.AttributeValue35);
                        ws.Cells[row, 121].PutValue(pi.AttributeName36);
                        ws.Cells[row, 122].PutValue(pi.AttributeValue36);
                        ws.Cells[row, 123].PutValue(pi.AttributeName37);
                        ws.Cells[row, 124].PutValue(pi.AttributeValue37);
                        ws.Cells[row, 125].PutValue(pi.AttributeName38);
                        ws.Cells[row, 126].PutValue(pi.AttributeValue38);
                        ws.Cells[row, 127].PutValue(pi.AttributeName39);
                        ws.Cells[row, 128].PutValue(pi.AttributeValue39);
                        ws.Cells[row, 129].PutValue(pi.AttributeName40);
                        ws.Cells[row, 130].PutValue(pi.AttributeValue40);
                        ws.Cells[row, 131].PutValue(pi.AttributeName41);
                        ws.Cells[row, 132].PutValue(pi.AttributeValue41);
                        ws.Cells[row, 133].PutValue(pi.AttributeName42);
                        ws.Cells[row, 134].PutValue(pi.AttributeValue42);
                        ws.Cells[row, 135].PutValue(pi.AttributeName43);
                        ws.Cells[row, 136].PutValue(pi.AttributeValue43);
                        ws.Cells[row, 137].PutValue(pi.AttributeName44);
                        ws.Cells[row, 138].PutValue(pi.AttributeValue44);
                        ws.Cells[row, 139].PutValue(pi.AttributeName45);
                        ws.Cells[row, 140].PutValue(pi.AttributeValue45);
                        ws.Cells[row, 141].PutValue(pi.AttributeName46);
                        ws.Cells[row, 142].PutValue(pi.AttributeValue46);
                        ws.Cells[row, 143].PutValue(pi.AttributeName47);
                        ws.Cells[row, 144].PutValue(pi.AttributeValue47);
                        ws.Cells[row, 145].PutValue(pi.AttributeName48);
                        ws.Cells[row, 146].PutValue(pi.AttributeValue48);
                        ws.Cells[row, 147].PutValue(pi.AttributeName49);
                        ws.Cells[row, 148].PutValue(pi.AttributeValue49);
                        ws.Cells[row, 149].PutValue(pi.AttributeName50);
                        ws.Cells[row, 150].PutValue(pi.AttributeValue50);
                        if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName1))
                            ws.Cells[row, 151].PutValue(pi.CustomColumnName1Value);
                        if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName2))
                            ws.Cells[row, 152].PutValue(pi.CustomColumnName2Value);
                        if (!string.IsNullOrEmpty(itemsList.FirstOrDefault().CustomColumnName3))
                            ws.Cells[row, 153].PutValue(pi.CustomColumnName3Value);
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
                        for (int col = 51; col <= ws.Cells.MaxColumn; col++)
                            ws.Cells[row, col].SetStyle(styleLeftAlignData);
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

        #region QC Item Update
        [HttpPost]
        [Route("QCItemUpdate")]
        public HttpResponseMessage QCItemUpdate([FromBody] QCItem qcItem)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                {
                    foreach (var modelError in ModelState.Values.SelectMany(v => v.Errors))
                    {
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, modelError.Exception.Message);
                    }
                }

                #region QC Item Attribute
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<QCItemAttribute>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, qcItem.ItemAttributes);

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
                    cmd.CommandText = "spQCItemUpdate";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@QCItemID", qcItem.QCItemID);
                    cmd.Parameters.AddWithValue("@ProductionItemID", qcItem.ProductionItemID);
                    cmd.Parameters.AddWithValue("@QCTestNo", qcItem.QCTestNo);
                    cmd.Parameters.AddWithValue("@NewShortDescription", qcItem.NewShortDescription);
                    cmd.Parameters.AddWithValue("@NewLongDescription", qcItem.NewLongDescription);
                    cmd.Parameters.AddWithValue("@MissingWords", qcItem.MissingWords);
                    cmd.Parameters.AddWithValue("@Noun", qcItem.Noun);
                    cmd.Parameters.AddWithValue("@Modifier", qcItem.Modifier);
                    cmd.Parameters.AddWithValue("@NounModifierComments", qcItem.NounModifierComments);
                    cmd.Parameters.AddWithValue("@QCStatus", qcItem.QCStatus);
                    cmd.Parameters.AddWithValue("@QCLevel", qcItem.QCLevel);
                    cmd.Parameters.AddWithValue("@QCLevelComments", qcItem.QCLevelComments);
                    cmd.Parameters.AddWithValue("@MFRName1", qcItem.MFRName1);
                    cmd.Parameters.AddWithValue("@MFRName1Comments", qcItem.MFRName1Comments);
                    cmd.Parameters.AddWithValue("@MFRPN1", qcItem.MFRPN1);
                    cmd.Parameters.AddWithValue("@MFRPN1Comments", qcItem.MFRPN1Comments);
                    cmd.Parameters.AddWithValue("@MFRName2", qcItem.MFRName2);
                    cmd.Parameters.AddWithValue("@MFRName2Comments", qcItem.MFRName2Comments);
                    cmd.Parameters.AddWithValue("@MFRPN2", qcItem.MFRPN2);
                    cmd.Parameters.AddWithValue("@MFRPN2Comments", qcItem.MFRPN2Comments);
                    cmd.Parameters.AddWithValue("@MFRName3", qcItem.MFRName3);
                    cmd.Parameters.AddWithValue("@MFRName3Comments", qcItem.MFRName3Comments);
                    cmd.Parameters.AddWithValue("@MFRPN3", qcItem.MFRPN3);
                    cmd.Parameters.AddWithValue("@MFRPN3Comments", qcItem.MFRPN3Comments);
                    cmd.Parameters.AddWithValue("@VendorName1", qcItem.VendorName1);
                    cmd.Parameters.AddWithValue("@VendorName1Comments", qcItem.VendorName1Comments);
                    cmd.Parameters.AddWithValue("@VendorPN1", qcItem.VendorPN1);
                    cmd.Parameters.AddWithValue("@VendorPN1Comments", qcItem.VendorPN1Comments);
                    cmd.Parameters.AddWithValue("@VendorName2", qcItem.VendorName2);
                    cmd.Parameters.AddWithValue("@VendorName2Comments", qcItem.VendorName2Comments);
                    cmd.Parameters.AddWithValue("@VendorPN2", qcItem.VendorPN2);
                    cmd.Parameters.AddWithValue("@VendorPN2Comments", qcItem.VendorPN2Comments);
                    cmd.Parameters.AddWithValue("@VendorName3", qcItem.VendorName3);
                    cmd.Parameters.AddWithValue("@VendorName3Comments", qcItem.VendorName3Comments);
                    cmd.Parameters.AddWithValue("@VendorPN3", qcItem.VendorPN3);
                    cmd.Parameters.AddWithValue("@VendorPN3Comments", qcItem.VendorPN3Comments);
                    cmd.Parameters.AddWithValue("@AdditionalInfoFromWeb", qcItem.AdditionalInfoFromWeb);
                    cmd.Parameters.AddWithValue("@AdditionalInfoFromWebComments", qcItem.AdditionalInfoFromWebComments);
                    cmd.Parameters.AddWithValue("@AdditionalInfoFromInput", qcItem.AdditionalInfoInput);
                    cmd.Parameters.AddWithValue("@AdditionalInfoFromInputComments", qcItem.AdditionalInfoInputComments);
                    cmd.Parameters.AddWithValue("@UNSPSCCode", qcItem.UNSPSCCode);
                    cmd.Parameters.AddWithValue("@UNSPSCCategory", qcItem.UNSPSCCategory);
                    cmd.Parameters.AddWithValue("@UNSPSCComments", qcItem.UNSPSCComments);
                    cmd.Parameters.AddWithValue("@WebRefURL1", qcItem.WebRefURL1);
                    cmd.Parameters.AddWithValue("@WebRefURL1Comments", qcItem.WebRefURL1Comments);
                    cmd.Parameters.AddWithValue("@WebRefURL2", qcItem.WebRefURL2);
                    cmd.Parameters.AddWithValue("@WebRefURL2Comments", qcItem.WebRefURL2Comments);
                    cmd.Parameters.AddWithValue("@WebRefURL3", qcItem.WebRefURL3);
                    cmd.Parameters.AddWithValue("@WebRefURL3Comments", qcItem.WebRefURL3Comments);
                    cmd.Parameters.AddWithValue("@PDFURL", qcItem.PDFURL);
                    cmd.Parameters.AddWithValue("@PDFURLComments", qcItem.PDFURLComments);
                    cmd.Parameters.AddWithValue("@Remarks", qcItem.Remarks);
                    cmd.Parameters.AddWithValue("@RemarksComments", qcItem.RemarksComments);
                    cmd.Parameters.AddWithValue("@Application", qcItem.Application);
                    cmd.Parameters.AddWithValue("@ApplicationComments", qcItem.ApplicationComments);
                    cmd.Parameters.AddWithValue("@DWG", qcItem.DWG);
                    cmd.Parameters.AddWithValue("@DWGComments", qcItem.DWGComments);
                    cmd.Parameters.AddWithValue("@ItemNo", qcItem.ItemNo);
                    cmd.Parameters.AddWithValue("@ItemNoComments", qcItem.ItemNoComments);
                    cmd.Parameters.AddWithValue("@OtherNo", qcItem.OtherNo);
                    cmd.Parameters.AddWithValue("@OtherNoComments", qcItem.OtherNoComments);
                    cmd.Parameters.AddWithValue("@POS", qcItem.POS);
                    cmd.Parameters.AddWithValue("@POSComments", qcItem.POSComments);
                    cmd.Parameters.AddWithValue("@SerialNo", qcItem.SerialNo);
                    cmd.Parameters.AddWithValue("@SerialNoComments", qcItem.SerialNoComments);
                    cmd.Parameters.AddWithValue("@KKSCode", qcItem.KKSCode);
                    cmd.Parameters.AddWithValue("@KKSCodeComments", qcItem.KKSCodeComments);
                    cmd.Parameters.AddWithValue("@AssemblyOrPart", qcItem.AssemblyOrPart);
                    cmd.Parameters.AddWithValue("@AssemblyOrPartComments", qcItem.AssemblyOrPartComments);
                    cmd.Parameters.AddWithValue("@BOM", qcItem.BOM);
                    cmd.Parameters.AddWithValue("@BOMComments", qcItem.BOMComments);
                    cmd.Parameters.AddWithValue("@GreenItems", qcItem.GreenItems);
                    cmd.Parameters.AddWithValue("@GreenItemsComments", qcItem.GreenItemsComments);
                    cmd.Parameters.Add(new SqlParameter("@AttributeNameValueAndComments", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", qcItem.UserID);
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

        #region QC Item Details
        [HttpGet]
        [Route("QCItemDetails/{QCItemID}")]
        public IHttpActionResult QCItemDetails(long QCItemID)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Item Attributes
                List<QCItemAttribute> QCItemAttributesList = new List<QCItemAttribute>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    #region QC Item Attributes
                    //Initialize command object
                    SqlCommand cmdItemAttributes = new SqlCommand();
                    cmdItemAttributes.Connection = conn;
                    cmdItemAttributes.CommandType = CommandType.StoredProcedure;
                    cmdItemAttributes.CommandText = "spQCItemAttributeDetails";

                    #region Adding Stored Procedure Parameters
                    cmdItemAttributes.Parameters.AddWithValue("@QCItemID", QCItemID);
                    #endregion

                    //Call sp to get all Item Attributes
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmdItemAttributes.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCItemAttribute qcItemAttribute = new QCItemAttribute();
                        qcItemAttribute.AttributeName = sqlReader["AttributeName"].ToString();
                        qcItemAttribute.AttributeValue = sqlReader["AttributeValue"].ToString();
                        qcItemAttribute.QCAttributeValue = sqlReader["QCAttributeValue"].ToString();
                        qcItemAttribute.QCAttributeValueComments = sqlReader["QCAttributeValueComments"].ToString();
                        QCItemAttributesList.Add(qcItemAttribute);
                    }
                    conn.Close();
                    #endregion

                    #region QC Item
                    QCItem qcItem = new QCItem();

                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCItemDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@QCItemID", QCItemID);

                    //Call sp to get all QC Item details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        qcItem.QCItemID = Convert.ToInt64(sqlReader["QCItemID"]);
                        qcItem.ProductionItemID = Convert.ToInt64(sqlReader["ProductionItemID"]);
                        qcItem.QCTestNo = Convert.ToInt32(sqlReader["QCTestNo"]);
                        qcItem.NewShortDescription = sqlReader["NewShortDescription"].ToString();
                        qcItem.NewLongDescription = sqlReader["NewLongDescription"].ToString();
                        qcItem.MissingWords = sqlReader["MissingWords"].ToString();
                        qcItem.Noun = sqlReader["Noun"].ToString();
                        qcItem.Modifier = sqlReader["Modifier"].ToString();
                        qcItem.NounModifierComments = sqlReader["NounModifierComments"].ToString();
                        qcItem.QCStatus = sqlReader["QCStatus"].ToString();
                        qcItem.QCLevel = sqlReader["QCLevel"].ToString();
                        qcItem.QCLevelComments = sqlReader["QCLevelComments"].ToString();
                        qcItem.MFRName1 = sqlReader["MFRName1"].ToString();
                        qcItem.MFRName1Comments = sqlReader["MFRName1Comments"].ToString();
                        qcItem.MFRPN1 = sqlReader["MFRPN1"].ToString();
                        qcItem.MFRPN1Comments = sqlReader["MFRPN1Comments"].ToString();
                        qcItem.MFRName2 = sqlReader["MFRName2"].ToString();
                        qcItem.MFRName2Comments = sqlReader["MFRName2Comments"].ToString();
                        qcItem.MFRPN2 = sqlReader["MFRPN2"].ToString();
                        qcItem.MFRPN2Comments = sqlReader["MFRPN2Comments"].ToString();
                        qcItem.MFRName3 = sqlReader["MFRName3"].ToString();
                        qcItem.MFRName3Comments = sqlReader["MFRName3Comments"].ToString();
                        qcItem.MFRPN3 = sqlReader["MFRPN3"].ToString();
                        qcItem.MFRPN3Comments = sqlReader["MFRPN3Comments"].ToString();
                        qcItem.VendorName1 = sqlReader["VendorName1"].ToString();
                        qcItem.VendorName1Comments = sqlReader["VendorName1Comments"].ToString();
                        qcItem.VendorPN1 = sqlReader["VendorPN1"].ToString();
                        qcItem.VendorPN1Comments = sqlReader["VendorPN1Comments"].ToString();
                        qcItem.VendorName2 = sqlReader["VendorName2"].ToString();
                        qcItem.VendorName2Comments = sqlReader["VendorName2Comments"].ToString();
                        qcItem.VendorPN2 = sqlReader["VendorPN2"].ToString();
                        qcItem.VendorPN2Comments = sqlReader["VendorPN2Comments"].ToString();
                        qcItem.VendorName3 = sqlReader["VendorName3"].ToString();
                        qcItem.VendorName3Comments = sqlReader["VendorName3Comments"].ToString();
                        qcItem.VendorPN3 = sqlReader["VendorPN3"].ToString();
                        qcItem.VendorPN3Comments = sqlReader["VendorPN3Comments"].ToString();
                        qcItem.AdditionalInfoFromWeb = sqlReader["AdditionalInfoFromWeb"].ToString();
                        qcItem.AdditionalInfoFromWebComments = sqlReader["AdditionalInfoFromWebComments"].ToString();
                        qcItem.AdditionalInfoInput = sqlReader["AdditionalInfoFromInput"].ToString();
                        qcItem.AdditionalInfoInputComments = sqlReader["AdditionalInfoFromInputComments"].ToString();
                        qcItem.UNSPSCVersion = sqlReader["UNSPSCVersion"].ToString();
                        qcItem.UNSPSCCode = sqlReader["UNSPSCCode"].ToString();
                        qcItem.UNSPSCCategory = sqlReader["UNSPSCCategory"].ToString();
                        qcItem.UNSPSCComments = sqlReader["UNSPSCComments"].ToString();
                        qcItem.WebRefURL1 = sqlReader["WebRefURL1"].ToString();
                        qcItem.WebRefURL1Comments = sqlReader["WebRefURL1Comments"].ToString();
                        qcItem.WebRefURL2 = sqlReader["WebRefURL2"].ToString();
                        qcItem.WebRefURL2Comments = sqlReader["WebRefURL2Comments"].ToString();
                        qcItem.WebRefURL3 = sqlReader["WebRefURL3"].ToString();
                        qcItem.WebRefURL3Comments = sqlReader["WebRefURL3Comments"].ToString();
                        qcItem.PDFURL = sqlReader["PDFURL"].ToString();
                        qcItem.PDFURLComments = sqlReader["PDFURLComments"].ToString();
                        qcItem.Remarks = sqlReader["Remarks"].ToString();
                        qcItem.RemarksComments = sqlReader["RemarksComments"].ToString();
                        qcItem.Application = sqlReader["Application"].ToString();
                        qcItem.ApplicationComments = sqlReader["ApplicationComments"].ToString();
                        qcItem.Application = sqlReader["Application"].ToString();
                        qcItem.DWG = sqlReader["DWG"].ToString();
                        qcItem.DWGComments = sqlReader["DWGComments"].ToString();
                        qcItem.ItemNo = sqlReader["ItemNo"].ToString();
                        qcItem.ItemNoComments = sqlReader["ItemNoComments"].ToString();
                        qcItem.OtherNo = sqlReader["OtherNo"].ToString();
                        qcItem.OtherNoComments = sqlReader["OtherNoComments"].ToString();
                        qcItem.POS = sqlReader["POS"].ToString();
                        qcItem.POSComments = sqlReader["POSComments"].ToString();
                        qcItem.SerialNo = sqlReader["SerialNo"].ToString();
                        qcItem.SerialNoComments = sqlReader["SerialNoComments"].ToString();
                        qcItem.KKSCode = sqlReader["KKSCode"].ToString();
                        qcItem.KKSCodeComments = sqlReader["KKSCodeComments"].ToString();
                        qcItem.AssemblyOrPart = sqlReader["AssemblyOrPart"].ToString();
                        qcItem.AssemblyOrPartComments = sqlReader["AssemblyOrPartComments"].ToString();
                        qcItem.BOM = sqlReader["BOM"].ToString();
                        qcItem.BOMComments = sqlReader["BOMComments"].ToString();
                        qcItem.GreenItems = sqlReader["GreenItems"].ToString();
                        qcItem.GreenItemsComments = sqlReader["GreenItemsComments"].ToString();
                        qcItem.QCUser = sqlReader["QCUser"].ToString();
                        qcItem.ItemAttributes = QCItemAttributesList;
                        conn.Close();
                    }

                    //return QC Item to the request
                    return Ok(qcItem);
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
    }
}
