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
using System.Data.OleDb;
using System.Data.Common;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/QCAllocation")]
    public class QCAllocationController : ApiController
    {
        #region Read Project Customer Codes
        [HttpGet]
        [Route("ReadProjectsCustomerCodes/{status?}")]
        public IHttpActionResult ReadProjectsCustomerCodes(string Status = "O")
        {
            try
            {
                //if (!AccessControl.CanUserAccessPage(ProductionUser, "QC Allocation"))
                //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                List<Customer> CustomerCodeList = new List<Customer>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCAllocationGetCustomerCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Status", Status);

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
                    cmd.CommandText = "spQCAllocationGetProjectCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);

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
                    cmd.CommandText = "spQCAllocationGetBatchNos";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);

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

        #region Read Project Or Batch Scope and Input Count
        [HttpGet]
        [Route("ReadProjectScope/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadProjectOrBatchScope(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                ProjectBatchScopeInputCount objProjectScope = new ProjectBatchScopeInputCount();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCAllocationProjectOrBatchScopeInputCount";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get scope and input count
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        objProjectScope.CustomerCode = sqlReader["CustomerCode"].ToString();
                        objProjectScope.ProjectCode = sqlReader["ProjectCode"].ToString();
                        objProjectScope.BatchNo = Convert.ToString(sqlReader["BatchNo"]);
                        objProjectScope.Scope = Convert.ToString(sqlReader["Scope"]);
                        objProjectScope.InputCount = Convert.ToInt64(sqlReader["InputCount"]);
                        objProjectScope.IsProjectAllocated = Convert.ToInt32(sqlReader["IsProjectAllocated"]);
                        objProjectScope.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        objProjectScope.QCCompletedPercentage = Convert.ToDecimal(sqlReader["QCCompletedPercentage"]);
                        conn.Close();

                        //return scope to the request
                        return Ok(objProjectScope);
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

        #region Validate and Allocate
        [HttpPost]
        [Route]
        public HttpResponseMessage ValidateAndAllocate([FromBody]QCAllocation qcAllocation)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid QC Allocation Data");

                if (!AccessControl.CanUserAccessPage(qcAllocation.UserID, "QC Allocation"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + qcAllocation.AllocatedFileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload QC Allocation file");
                #endregion

                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + qcAllocation.AllocatedFileName);
                string InputFileExtension = Path.GetExtension(InputFilepath);
                int ProductionUserColumnNo = -1, QCUserColumnNo = -1, ActivitiesColumnNo = -1, StatusColumnNo = -1, ProductionCommentsColumnNo=-1; //Unique Column will be checked in back end

                Workbook wbIF = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.LoadData(InputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file has data rows
                if (wsIF.Cells.MaxRow <= 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded QC Allocation file has no data rows.");
                #endregion

                #region Check uploaded file for duplicate columns if any
                for (int mcol = 0; mcol < wsIF.Cells.MaxColumn; mcol++)              //main column
                {
                    for (int ncol = mcol + 1; ncol <= wsIF.Cells.MaxColumn; ncol++)      //navigation column
                    {
                        if (wsIF.Cells[0, mcol].StringValue.Trim().ToLower() == wsIF.Cells[0, ncol].StringValue.Trim().ToLower())
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded QC Allocation file has duplicate columns. column name:" + wsIF.Cells[0, mcol].StringValue.Trim());
                    }
                }
                #endregion

                #region Check any value exceeds 4000 characters from all columns
                for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
                {
                    for (int row = 1; row <= wsIF.Cells.MaxRow; row++)
                    {
                        if (wsIF.Cells[row, col].StringValue.Length > 4000)
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'" + wsIF.Cells[0, col].StringValue + "' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Check Production User, QC User, Activities,Status,Production Comments columns exists, if exists get column Nos.
                for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[0, col].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Blank/Empty column heading in uploaded file. Column No.: " + (col + 1).ToString());

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "production user")
                        ProductionUserColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "qc user")
                        QCUserColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "activities")
                        ActivitiesColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "status")
                        StatusColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "production comments")
                        ProductionCommentsColumnNo = col;

                    if (ProductionUserColumnNo >= 0 && QCUserColumnNo >= 0 && ActivitiesColumnNo >= 0 && StatusColumnNo >= 0 && ProductionCommentsColumnNo >= 0)
                        break;
                }

                if (ProductionUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column not found in uploaded QC allocation file");

                if (QCUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'QC User' column not found in uploaded QC allocation file");

                if (ActivitiesColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column not found in uploaded QC allocation file");

                if (StatusColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Status' column not found in uploaded QC allocation file");

                if (ProductionCommentsColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production Comments' column not found in uploaded QC allocation file");
                #endregion

                #region Check empty/blank value exists in Production User, QC User, Activities columns, and status as 'Production Completed' value
                for (int row = 1; row <= wsIF.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[row, ProductionUserColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production user' column cannot contain blank/empty value");

                    if (wsIF.Cells[row, ProductionUserColumnNo].StringValue.Length > 50)
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column value cannot exceed 50 characters. Row No.:" + row.ToString());

                    if (string.IsNullOrEmpty(wsIF.Cells[row, QCUserColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'QC User' column cannot contain blank/empty value");

                    if (wsIF.Cells[row, QCUserColumnNo].StringValue.Length > 50)
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'QC User' column value cannot exceed 50 characters. Row No.:" + row.ToString());

                    if (string.IsNullOrEmpty(wsIF.Cells[row, ActivitiesColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column cannot contain blank/empty value");

                    //Activities length cannot exceed 4000 characters which is checked above for all columns

                    if (wsIF.Cells[row, StatusColumnNo].StringValue.Trim().ToLower() != "production completed")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "All the rows should have status as 'Production Completed'");
                }
                #endregion

                #region Form string to create temp table
                string sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                #region Generate dynamic column name and data type string
                string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
                {
                    if (!string.IsNullOrEmpty(wsIF.Cells[0, col].StringValue.Trim()))
                    {
                        ColumnName = "[" + wsIF.Cells[0, col].StringValue.Trim() + "]";
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

                #region Create temp table, write input file data, allocate QC, Move file
                #region Creating SQL table
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = strSQL;
                cmd.CommandType = System.Data.CommandType.Text;
                conn.Open();
                cmd.CommandTimeout = 0;
                cmd.ExecuteNonQuery();
                #endregion

                #region Writing data to SQL table
                DataFormatConverter dataFormatConverter = new DataFormatConverter();
                DataTable excelData = dataFormatConverter.ExcelToDataTable(InputFilepath);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion

                #region Create QC Allocation and Move file
                //Initialize command object
                SqlCommand cmdPA = new SqlCommand();
                cmdPA.Connection = conn;
                cmdPA.CommandType = CommandType.StoredProcedure;
                cmdPA.CommandText = "spQCAllocationValidateAndAllocate";

                #region Adding Stored Procedure Parameters
                cmdPA.Parameters.AddWithValue("@CustomerCode", qcAllocation.CustomerCode);
                cmdPA.Parameters.AddWithValue("@ProjectCode", qcAllocation.ProjectCode);
                cmdPA.Parameters.AddWithValue("@BatchNo", qcAllocation.BatchNo);
                cmdPA.Parameters.AddWithValue("@AllocatedFileExtension", InputFileExtension);
                cmdPA.Parameters.AddWithValue("@TempTableName", sqlTableName);
                cmdPA.Parameters.AddWithValue("@UserID", qcAllocation.UserID);
                #endregion

                //Calling sp to create QC Allocation
                cmdPA.CommandTimeout = 0;
                string Result = cmdPA.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split(',');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    long QCAllocationID = Convert.ToInt64(arrResult[1]);
                    string NewAllocatedFileName = arrResult[2];

                    #region Input File move Starts
                    if (File.Exists(dirTemp + qcAllocation.AllocatedFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/"));
                        FileOperations.MoveFile(dirTemp, qcAllocation.AllocatedFileName, dirUploads, NewAllocatedFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, QCAllocationID);
                }
                else
                    //return error response status code
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
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

        #region View Existing QC Allocation

        #region Download QC Completed Output Table
        [HttpGet]
        [Route("DownloadQCCompletedOutputTable/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public HttpResponseMessage DownloadQCCompletedOutputTable(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                string FileName = string.Empty;

                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                if (string.IsNullOrEmpty(BatchNo))
                    FileName = "QCCompletedOutputTable_" + CustomerCode + '_' + ProjectCode + ".xlsx";
                else
                    FileName = "QCCompletedOutputTable_" + CustomerCode + '_' + ProjectCode + '_' + BatchNo + ".xlsx";

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
                    cmd.CommandText = "spQCAllocationDownloadOutputTable";
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
                        //Writing column headings
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                            ws.Cells[0, col].PutValue(sqlReader.GetName(col));

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                                ws.Cells[row, col].PutValue(sqlReader.GetValue(col));
                            ++row;
                        }

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

        #region Read Existing Project QC Allocations
        [HttpGet]
        [Route("ReadExistingProjectQCAllocations/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadExistingProjectQCAllocations(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Existing QC Allocations
                List<QCExistingAllocation> qcExistingAllocationList = new List<QCExistingAllocation>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCAllocationGetExistingProjectAllocations";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of existing QC Allocations
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCExistingAllocation qcExistingAllocation = new QCExistingAllocation();

                        qcExistingAllocation.QCAllocationID = Convert.ToInt64(sqlReader["QCAllocationID"]);
                        qcExistingAllocation.AllocatedOn = Convert.ToDateTime(sqlReader["AllocatedOn"]);
                        qcExistingAllocation.AllocatedByUserName = sqlReader["AllocatedByUserName"].ToString();
                        qcExistingAllocation.AllocatedFileName = sqlReader["AllocatedFileName"].ToString();
                        qcExistingAllocation.AllocatedCount = Convert.ToInt32(sqlReader["AllocatedCount"]);
                        qcExistingAllocation.CompletedCount = Convert.ToInt32(sqlReader["CompletedCount"]);

                        qcExistingAllocationList.Add(qcExistingAllocation);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(qcExistingAllocationList);
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

        #region Download Allocation QC Completed All Details
        [HttpGet]
        [Route("DownloadAllocationQCCompletedAllDetails/{id}")]
        public HttpResponseMessage DownloadAllocationQCCompletedAllDetails(long id)
        {
            try
            {
                string FileName = "QCCompletedAllocationAllDetails_" + id.ToString() + ".xlsx";

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
                    cmd.CommandText = "spQCAllocationDownloadAllUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@QCAllocationID", id);

                    //Calling sp to get uploaded data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        //Writing column headings
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
                                ws.Cells[0, col].PutValue(sqlReader.GetName(col));
                        }

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
                                    ws.Cells[row, col].PutValue(sqlReader.GetValue(col));
                            }
                            ++row;
                        }

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

        #region Read Existing Project Allocation Details By QC Allocation ID
        [HttpGet]
        [Route("ReadExistingProjectAllocationDetailsByID/{id}")]
        public IHttpActionResult ReadExistingProjectAllocationDetailsByID(long id)
        {
            try
            {
                int index = 0;  //temporary added since bootstrap table requires a key to render table
                //Create a list to hold the list of Existing QC Allocations
                List<QCAllocationDetails> qcAllocationDetailsList = new List<QCAllocationDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCAllocationGetExistingDetailsByID";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCAllocationID", id);

                    //Calling sp to get list of existing QC Allocation details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        QCAllocationDetails qcAllocationDetails = new QCAllocationDetails();

                        qcAllocationDetails.QCAllocationDetailsID = index;
                        qcAllocationDetails.QCAllocationID = Convert.ToInt64(sqlReader["QCAllocationID"]);
                        qcAllocationDetails.Activities = sqlReader["Activities"].ToString();
                        qcAllocationDetails.QCUser = sqlReader["QCUser"].ToString();
                        qcAllocationDetails.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                        qcAllocationDetails.QCPendingCount = Convert.ToInt32(sqlReader["QCPendingCount"]);
                        qcAllocationDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);

                        qcAllocationDetailsList.Add(qcAllocationDetails);
                        index++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(qcAllocationDetailsList);
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

        #region Download Allocated file by Project Allocation ID
        [HttpGet]
        [Route("downloadfile/{id}")]
        public HttpResponseMessage DownloadAllocatedFile(long id)
        {
            try
            {
                string AllocatedFileName = string.Empty;
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/");
                string fileFullName = string.Empty;

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spQCAllocationGetUploadedFileNameByID";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCAllocationID", id);

                    //Calling sp to get Allocated FileName
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        AllocatedFileName = sqlReader["AllocatedFileName"].ToString();
                        fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        //Check whether File exists.
                        if (!File.Exists(fileFullName))
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocated File not found");

                        //Read the file into a Byte Array.
                        byte[] bytes = File.ReadAllBytes(fileFullName);

                        //Set the Response Content.
                        response.Content = new ByteArrayContent(bytes);

                        //Set the Response Content Length.
                        response.Content.Headers.ContentLength = bytes.LongLength;

                        //Set the Content Disposition Header Value and FileName.
                        response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                        response.Content.Headers.ContentDisposition.FileName = AllocatedFileName;

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

        #region Delete QC Allocation
        [HttpPatch]
        [Route("DeleteQCAllocation/{id}/{UserID}")]
        public HttpResponseMessage DeleteQCAllocation(long id, string UserID)
        {
            try
            {
                string AllocatedFileName = string.Empty, fileFullName = string.Empty;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spDeleteQCAllocation";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCAllocationID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete QC Allocation
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/");
                        fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        if (Convert.ToInt16(arrResult[1]) > 0)
                        {
                            #region Delete Allocated File
                            if (File.Exists(fileFullName))
                                File.Delete(fileFullName);
                            #endregion

                            return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                        }
                        else
                        {
                            #region Deleting rows from allocated file
                            int UniqueColumnNameColNo = -1;
                            Workbook wbAF = new Workbook();         //Allocated file
                            Aspose.Cells.License l = new Aspose.Cells.License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the QC uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> QCUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spQCAllocationGetAllUploadedSKUsById";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@QCAllocationID", id);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                QCUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                if (!QCUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
                                {
                                    ws.Cells.DeleteRow(sRow);
                                    maxRows = ws.Cells.MaxRow;
                                    if (sRow > 1)
                                        sRow = sRow - 1;
                                }
                                else
                                    sRow++;
                            }

                            wbAF.Save(fileFullName);
                            #endregion

                            return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                        }
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

        #region QC Allocation Details
        #region Download QC Completed Allocation Activities of user
        [HttpPost]
        [Route("DownloadQCCompletedAllocationActivities")]
        public HttpResponseMessage DownloadQCCompletedAllocationActivities([FromBody] QCDownload qcDownload)
        {
            try
            {
                string FileName = "QCCompletedActivities_" + qcDownload.QCUser + '_' + qcDownload.QCAllocationID.ToString() + ".xlsx";

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
                    cmd.CommandText = "spQCAllocationDownloadActivitiesUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@QCAllocationID", qcDownload.QCAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", qcDownload.Activities);
                    cmd.Parameters.AddWithValue("@QCUser", qcDownload.QCUser);

                    //Calling sp to get uploaded data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        //Writing column headings
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
                                ws.Cells[0, col].PutValue(sqlReader.GetName(col));
                        }

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
                                    ws.Cells[row, col].PutValue(sqlReader.GetValue(col));
                            }
                            ++row;
                        }

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

        #region Delete QC Allocation Activities
        [HttpPatch]
        [Route("DeleteQCAllocationActivities")]
        public HttpResponseMessage DeleteQCAllocationActivities([FromBody] QCAllocationDetails qcAllocationDetails)
        {
            try
            {
                string AllocatedFileName = string.Empty, fileFullName = string.Empty;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spDeleteQCAllocationUserActivities";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCAllocationID", qcAllocationDetails.QCAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", qcAllocationDetails.Activities);
                    cmd.Parameters.AddWithValue("@QCUser", qcAllocationDetails.QCUser);
                    cmd.Parameters.AddWithValue("@UserID", qcAllocationDetails.UserID);

                    //Calling sp to delete QC Allocation Activities
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/");
                        fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        if (Convert.ToInt16(arrResult[1]) > 0)
                        {
                            #region Delete Allocated File
                            if (File.Exists(fileFullName))
                                File.Delete(fileFullName);
                            #endregion

                            return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                        }
                        else
                        {
                            #region Deleting rows from allocated file
                            int QCUserColumnNo = -1, ActivitiesColNo = -1, UniqueColumnNameColNo = -1;
                            Workbook wbAF = new Workbook();         //Allocated file
                            Aspose.Cells.License l = new Aspose.Cells.License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "qc user")
                                    QCUserColumnNo = sCol;

                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "activities")
                                    ActivitiesColNo = sCol;

                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (QCUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the QC uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> QCUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spQCAllocationGetAllQCUploadedSKUs";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@QCAllocationID", qcAllocationDetails.QCAllocationID);
                            cmdPU.Parameters.AddWithValue("@Activities", qcAllocationDetails.Activities);
                            cmdPU.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                QCUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                if (ws.Cells[sRow, QCUserColumnNo].StringValue.Trim().ToLower() == qcAllocationDetails.QCUser.Trim().ToLower() &&
                                    ws.Cells[sRow, ActivitiesColNo].StringValue.Trim().ToLower() == qcAllocationDetails.Activities.Trim().ToLower())
                                {
                                    UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                    if (!QCUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
                                    {
                                        ws.Cells.DeleteRow(sRow);
                                        maxRows = ws.Cells.MaxRow;
                                        if (sRow > 1)
                                            sRow = sRow - 1;
                                    }
                                    else
                                        sRow++;
                                }
                                else
                                    sRow++;
                            }

                            wbAF.Save(fileFullName);
                            #endregion

                            return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                        }
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

        #region Change QC Allocation User
        [HttpPatch]
        [Route("ChangeUser")]
        public HttpResponseMessage ChangeQCAllocationUser([FromBody] QCAllocationDetails qcAllocationDetails)
        {
            try
            {
                string ChangeToQCUser = string.Empty;

                if (qcAllocationDetails.ChangeToQCUser.Contains('-'))
                {
                    string[] arrChangeToQCUser = qcAllocationDetails.ChangeToQCUser.Split('-');
                    ChangeToQCUser = arrChangeToQCUser[1].Trim();
                }
                else
                    ChangeToQCUser = qcAllocationDetails.ChangeToQCUser;

                if (string.IsNullOrEmpty(qcAllocationDetails.Activities))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid activities");

                if (qcAllocationDetails.QCUser.Trim().ToLower() == ChangeToQCUser.Trim().ToLower())
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "QC User and Change To QC User cannot be same.");

                if (string.IsNullOrEmpty(qcAllocationDetails.UserID))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid User");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spQCAllocationChangeQCUser";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@QCAllocationID", qcAllocationDetails.QCAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", qcAllocationDetails.Activities);
                    cmd.Parameters.AddWithValue("@QCUser", qcAllocationDetails.QCUser);
                    cmd.Parameters.AddWithValue("@ChangeToQCUser", ChangeToQCUser);
                    cmd.Parameters.AddWithValue("@UserID", qcAllocationDetails.UserID);

                    //Calling sp to change the QC Allocation User
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "updated")
                    {
                        #region Changing user in file
                        string AllocatedFileName = arrResult[1];
                        string UniqueColumnName = arrResult[2];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/QCAllocation/");
                        string fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        int QCUserColumnNo = -1, ActivitiesColNo = -1, UniqueColumnNameColNo = -1;
                        Workbook wbAF = new Workbook();         //Allocated file
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wbAF.LoadData(fileFullName);
                        var ws = wbAF.Worksheets[0];

                        for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                        {
                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "qc user")
                                QCUserColumnNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "activities")
                                ActivitiesColNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                UniqueColumnNameColNo = sCol;

                            if (QCUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                                break;
                        }

                        if (QCUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                        {
                            #region Fetch all the QC uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> QCUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spQCAllocationGetAllQCUploadedSKUs";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@QCAllocationID", qcAllocationDetails.QCAllocationID);
                            cmdPU.Parameters.AddWithValue("@Activities", qcAllocationDetails.Activities);
                            cmdPU.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                QCUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            for (int sRow = 1; sRow <= ws.Cells.MaxRow; sRow++)
                            {
                                if (ws.Cells[sRow, QCUserColumnNo].StringValue.Trim().ToLower() == qcAllocationDetails.QCUser.Trim().ToLower() &&
                                        ws.Cells[sRow, ActivitiesColNo].StringValue.Trim().ToLower() == qcAllocationDetails.Activities.Trim().ToLower())
                                {
                                    UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                    if (!QCUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
                                        ws.Cells[sRow, QCUserColumnNo].PutValue(ChangeToQCUser);
                                }
                            }
                            wbAF.Save(fileFullName);
                        }
                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, Result.Trim().ToLower());
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
        #endregion

        #endregion
    }
}
