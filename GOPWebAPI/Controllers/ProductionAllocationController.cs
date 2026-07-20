using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
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
    [RoutePrefix("api/ProductionAllocation")]
    public class ProductionAllocationController : ApiController
    {
        #region Read OnGoing Projects Customer Codes
        [HttpGet]
        [Route]
        public IHttpActionResult ReadOnGoingProjectsCustomerCodes()
        {
            try
            {
                //if (!AccessControl.CanUserAccessPage(UserID, "Production Allocation"))
                //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                List<Customer> CustomerCodeList = new List<Customer>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectsCustomerCodes";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of projects
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
                    return Ok(CustomerCodeList.Distinct());
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

        #region Read OnGoing Project Codes of Customer
        [HttpGet]
        [Route("ReadOnGoingProjectCodesOfCustomer/{CustomerCode}")]
        public IHttpActionResult ReadOnGoingProjectCodesOfCustomer(string CustomerCode)
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> ProjectCodeList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectCodesOfCustomer";
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

        #region Read OnGoing Batches of Project
        [HttpGet]
        [Route("ReadOnGoingBatchesOfProject/{CustomerCode}/{ProjectCode}")]
        public IHttpActionResult ReadOnGoingBatchesOfProject(string CustomerCode, string ProjectCode)
        {
            try
            {
                List<CustomerCodeProjectCodeBatchNo> BatchNoList = new List<CustomerCodeProjectCodeBatchNo>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingBatchesOfProject";
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
                    cmd.CommandText = "spProductionAllocationProjectOrBatchScopeInputCount";
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
                        objProjectScope.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        objProjectScope.ProductionCompletedPercentage = Convert.ToDecimal(sqlReader["ProductionCompletedPercentage"]);
                        objProjectScope.IsProjectSettingsExist = Convert.ToBoolean(sqlReader["IsProjectSettingsExist"]);
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

        #region Read Project Activities for Excel Based Upload Projects
        [HttpGet]
        [Route("ReadProjectActivities/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadProjectActivities(string CustomerCode, string ProjectCode, string BatchNo = "")
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
                    cmd.CommandText = "spProductionAllocationProjectActivityCountStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of Activity Counts
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectActivitiesCount pac = new ProjectActivitiesCount();

                        pac.Activity = sqlReader["Activity"].ToString();
                        pac.ActivityCount = Convert.ToInt32(sqlReader["ActivityCount"]);
                        pac.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        pac.ProductionPendingToAllocate = Convert.ToInt32(sqlReader["ProductionPendingToAllocate"]);
                        pac.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        pac.ProductionCompletionPercentage = Convert.ToDecimal(sqlReader["ProductionCompletionPercentage"]);
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

        #region Read Project Allocation and Completed Count Status for Onscreen editing projects
        [HttpGet]
        [Route("ReadProjectAllocationAndCompletedCountStatus")]
        public IHttpActionResult ReadProjectAllocationAndCompletedCountStatus(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                ProjectAllocationAndCompletedCountStatus projectAllocationAndCompletedCountStatus = new ProjectAllocationAndCompletedCountStatus();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationProjectCountStatusForOnScreenEditingProjects";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get Count status
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        projectAllocationAndCompletedCountStatus.Activities = sqlReader["Activities"].ToString();
                        projectAllocationAndCompletedCountStatus.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        projectAllocationAndCompletedCountStatus.ProductionPendingToAllocate = Convert.ToInt32(sqlReader["ProductionPendingToAllocate"]);
                        projectAllocationAndCompletedCountStatus.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        projectAllocationAndCompletedCountStatus.ProductionCompletionPercentage = Convert.ToDecimal(sqlReader["ProductionCompletionPercentage"]);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectAllocationAndCompletedCountStatus);
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

        #region Read Unique Column Names
        [HttpGet]
        [Route("ReadUniqueColumnNames/{FileName}/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadUniqueColumnNames(string FileName, string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            //Swagger fix
            if (BatchNo == "{BatchNo}")
                BatchNo = "";

            //Create a list to hold the list of Unique Column Names
            List<UniqueColumn> UniqueColumnNamesList = new List<UniqueColumn>();
            System.Data.Common.DbDataReader sqlReader;

            using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
            {
                //Initialize command object
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "spProductionAllocationUniqueColumnName";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                //Add parameters with values
                cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                #region Return the unique column name of last production allocation
                conn.Open();
                cmd.CommandTimeout = 0;
                sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                if (sqlReader.Read())
                {
                    UniqueColumn uniqueColumn = new UniqueColumn();

                    uniqueColumn.UniqueColumnName = sqlReader["UniqueColumnName"].ToString();
                    conn.Close();

                    UniqueColumnNamesList.Add(uniqueColumn);

                    //return list to the request
                    return Ok(UniqueColumnNamesList);
                }
                #endregion

                #region Read unique Column Names from file
                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                Workbook IFwb = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //IFwb.Open(InputFilepath);
                IFwb.LoadData(InputFilepath);
                Worksheet IFws = IFwb.Worksheets[0];

                for (int col = 0; col <= IFws.Cells.MaxColumn; col++)
                {
                    UniqueColumn uniqueColumn = new UniqueColumn();
                    uniqueColumn.UniqueColumnName = IFws.Cells[0, col].StringValue.Trim();
                    if (!UniqueColumnNamesList.Any(u => u.UniqueColumnName.Trim().ToLower() == uniqueColumn.UniqueColumnName.ToLower()))
                        UniqueColumnNamesList.Add(uniqueColumn);
                }

                return Ok(UniqueColumnNamesList);
                #endregion
            }

        }
        #endregion

        #region Validate and Allocate
        [HttpPost]
        [Route]
        public HttpResponseMessage ValidateAndAllocate([FromBody] ProductionAllocation productionAllocation)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Production Allocation Data");

                if (!AccessControl.CanUserAccessPage(productionAllocation.UserID, "Production Allocation"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + productionAllocation.AllocatedFileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload Production Allocation file");
                #endregion

                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + productionAllocation.AllocatedFileName);
                string InputFileExtension = Path.GetExtension(InputFilepath);
                int UniqueColumnNameColumnNo = -1, ProductionUserColumnNo = -1, ActivitiesColumnNo = -1, StatusColumnNo = -1;

                Workbook wbIF = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wbIF.Open(InputFilepath);
                wbIF.LoadData(InputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file has data rows
                if (wsIF.Cells.MaxRow <= 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Allocation file has no data rows.");
                #endregion

                #region Check uploaded file for duplicate columns if any
                for (int mcol = 0; mcol < wsIF.Cells.MaxColumn; mcol++)              //main column
                {
                    for (int ncol = mcol + 1; ncol <= wsIF.Cells.MaxColumn; ncol++)      //navigation column
                    {
                        if (wsIF.Cells[0, mcol].StringValue.Trim().ToLower() == wsIF.Cells[0, ncol].StringValue.Trim().ToLower())
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded Production Allocation file has duplicate columns. column name:" + wsIF.Cells[0, mcol].StringValue.Trim());
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

                #region Check Unique Column Name,Production User, Activities, Status columns exists, if exists get column Nos.
                for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[0, col].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Blank/Empty column heading in allocation file. Column No.: " + (col + 1).ToString());

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == productionAllocation.UniqueColumnName.Trim().ToLower())
                        UniqueColumnNameColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "activities")
                        ActivitiesColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "production user")
                        ProductionUserColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == "status")
                        StatusColumnNo = col;

                    if (UniqueColumnNameColumnNo >= 0 && ActivitiesColumnNo >= 0 && ProductionUserColumnNo >= 0 && StatusColumnNo >= 0)
                        break;
                }

                if (UniqueColumnNameColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'" + productionAllocation.UniqueColumnName + "' column not found in uploaded production allocation file");

                if (ActivitiesColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column not found in uploaded production allocation file");

                if (ProductionUserColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column not found in uploaded production allocation file");
                #endregion

                #region Check empty/blank value exists in Unique Column Name, Activities, Production User columns and character length
                for (int row = 1; row <= wsIF.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[row, UniqueColumnNameColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'" + productionAllocation.UniqueColumnName + "' column cannot contain blank/empty value. Row No.:" + row.ToString());

                    if (wsIF.Cells[row, UniqueColumnNameColumnNo].StringValue.Length > 100)
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'" + productionAllocation.UniqueColumnName + "' column value cannot exceed 100 characters. Row No.:" + row.ToString());

                    if (string.IsNullOrEmpty(wsIF.Cells[row, ActivitiesColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Activities' column cannot contain blank/empty value, check if ");

                    //Activities length cannot exceed 4000 characters which is checked above for all columns

                    if (string.IsNullOrEmpty(wsIF.Cells[row, ProductionUserColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column cannot contain blank/empty value");

                    if (wsIF.Cells[row, ProductionUserColumnNo].StringValue.Length > 50)
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Production User' column value cannot exceed 50 characters. Row No.:" + row.ToString());
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

                if (StatusColumnNo < 0)
                    ColumnNameDataTypeString += ",Status VARCHAR(50)";

                strSQL += ColumnNameDataTypeString + ")";
                #endregion

                #region Create temp table, write input file data, allocate Production, Move file
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

                #region Create Production Allocation
                //Initialize command object
                SqlCommand cmdPA = new SqlCommand();
                cmdPA.Connection = conn;
                cmdPA.CommandType = CommandType.StoredProcedure;
                cmdPA.CommandText = "spProductionAllocationValidateAndAllocate";

                #region Adding Stored Procedure Parameters
                cmdPA.Parameters.AddWithValue("@CustomerCode", productionAllocation.CustomerCode);
                cmdPA.Parameters.AddWithValue("@ProjectCode", productionAllocation.ProjectCode);
                cmdPA.Parameters.AddWithValue("@BatchNo", productionAllocation.BatchNo);
                cmdPA.Parameters.AddWithValue("@AllocatedFileExtension", InputFileExtension);
                cmdPA.Parameters.AddWithValue("@UniqueColumnName", productionAllocation.UniqueColumnName);
                cmdPA.Parameters.AddWithValue("@TempTableName", sqlTableName);
                cmdPA.Parameters.AddWithValue("@UserID", productionAllocation.UserID);
                #endregion

                //Calling sp to create Production Allocation
                cmdPA.CommandTimeout = 0;
                string Result = cmdPA.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split(',');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    long ProductionAllocationID = Convert.ToInt64(arrResult[1]);
                    string NewAllocatedFileName = arrResult[2];

                    #region Input File move Starts
                    if (File.Exists(dirTemp + productionAllocation.AllocatedFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/"));
                        FileOperations.MoveFile(dirTemp, productionAllocation.AllocatedFileName, dirUploads, NewAllocatedFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, ProductionAllocationID);
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

        #region View Existing Production Allocation

        #region Download Production Completed Output Table
        [HttpGet]
        [Route("DownloadProductionCompletedOutputTable/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public HttpResponseMessage DownloadProductionCompletedOutputTable(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                string FileName = string.Empty;

                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                if (string.IsNullOrEmpty(BatchNo))
                    FileName = "ProductionCompletedOutputTable_" + CustomerCode + '_' + ProjectCode + ".xlsx";
                else
                    FileName = "ProductionCompletedOutputTable_" + CustomerCode + '_' + ProjectCode + '_' + BatchNo + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationDownloadOutputTable";
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

        #region Read Existing Project Allocations for both Excel Based Uploads and On Screen Editing Projects
        [HttpGet]
        [Route("ReadExistingProjectAllocations/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadExistingProjectAllocations(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                //Create a list to hold the list of Existing Production Allocations
                List<ProductionExistingAllocation> productionExistingAllocationList = new List<ProductionExistingAllocation>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationGetExistingProjectAllocations";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of existing Production Allocations
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionExistingAllocation productionExistingAllocation = new ProductionExistingAllocation();

                        productionExistingAllocation.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionExistingAllocation.AllocatedOn = Convert.ToDateTime(sqlReader["AllocatedOn"]);
                        productionExistingAllocation.AllocatedByUserName = sqlReader["AllocatedByUserName"].ToString();
                        productionExistingAllocation.AllocatedFileName = sqlReader["AllocatedFileName"].ToString();
                        productionExistingAllocation.AllocatedCount = Convert.ToInt32(sqlReader["AllocatedCount"]);
                        productionExistingAllocation.CompletedCount = Convert.ToInt32(sqlReader["CompletedCount"]);

                        productionExistingAllocationList.Add(productionExistingAllocation);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(productionExistingAllocationList);
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

        #region Download Allocated file by Production Allocation ID Same for both Excel based and On Screen
        [HttpGet]
        [Route("downloadfile/{id}")]
        public HttpResponseMessage DownloadAllocatedFile(long id)
        {
            try
            {
                string AllocatedFileName = string.Empty;
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
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
                    cmd.CommandText = "spProductionAllocationGetUploadedFileNameByID";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);

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

        #region Download Production Allocation Production Completed All Details for Excel Based Uploads
        [HttpGet]
        [Route("DownloadAllocationProductionCompletedAllDetails/{id}")]
        public HttpResponseMessage DownloadAllocationProductionCompletedAllDetails(long id)
        {
            try
            {
                string FileName = "ProductionCompletedAllocationAllDetails_" + id.ToString() + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationDownloadAllUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);

                    //Calling sp to get uploaded data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        //Writing column headings
                        for (int col = 0; col < sqlReader.FieldCount; ++col)
                        {
                            //if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
                            ws.Cells[0, col].PutValue(sqlReader.GetName(col));
                        }

                        //Writing data rows
                        while (sqlReader.Read())
                        {
                            for (int col = 0; col < sqlReader.FieldCount; ++col)
                            {
                                //if (sqlReader.GetName(col).Trim().ToLower() != "isdeleted")
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

        #region Download Production Allocation Production Completed All Details for On Screen Editing Projects 
        [HttpGet]
        [Route("DownloadAllocationProductionCompletedSKUs")]
        public HttpResponseMessage DownloadAllocationProductionCompletedSKUs(long id)
        {
            try
            {
                string FileName = "Production Completed All SKUs of Production Allocation ID - " + id.ToString() + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationDownloadProductionCompletedSKUsForOnScreenEditingProjects";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);

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

        #region Read Existing Project Allocation Details By Production Allocation ID for both Excel Based Uploads and On Screen Editing Projects
        [HttpGet]
        [Route("ReadExistingProjectAllocationDetailsByID/{id}")]
        public IHttpActionResult ReadExistingProjectAllocationDetailsByID(long id)
        {
            try
            {
                int index = 0;  //temporary added since bootstrap table requires a key to render table
                //Create a list to hold the list of Existing Production Allocations
                List<ProductionAllocationDetails> productionAllocationDetailsList = new List<ProductionAllocationDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationGetExistingDetailsByID";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);

                    //Calling sp to get list of existing Production Allocation details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionAllocationDetails productionAllocationDetails = new ProductionAllocationDetails();

                        productionAllocationDetails.ProductionAllocationDetailsID = index;
                        productionAllocationDetails.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionAllocationDetails.Activities = sqlReader["Activities"].ToString();
                        productionAllocationDetails.ProductionUser = sqlReader["ProductionUser"].ToString();
                        productionAllocationDetails.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        productionAllocationDetails.ProductionPendingCount = Convert.ToInt32(sqlReader["ProductionPendingCount"]);
                        productionAllocationDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);

                        productionAllocationDetailsList.Add(productionAllocationDetails);
                        index++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(productionAllocationDetailsList);
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

        #region Delete Production Allocation Production Pending SKUs for Excel Based Uploads
        [HttpPatch]
        [Route("DeleteProductionAllocation/{id}/{UserID}")]
        public HttpResponseMessage DeleteProductionAllocation(long id, string UserID)
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
                    cmd.CommandText = "spDeleteProductionAllocation";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete Production Allocation
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
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
                            //wbAF.Open(fileFullName);
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the production uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> ProductionUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spProductionAllocationGetAllUploadedSKUsById";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@ProductionAllocationID", id);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                ProductionUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                if (!ProductionUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
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

        #region Delete Production Allocation Production Pending SKUs for Screen Editing Based Projects
        [HttpPatch]
        [Route("DeleteProductionAllocationProductionPendingSKUs")]
        public HttpResponseMessage DeleteProductionAllocationProductionPendingSKUs(long id, string UserID)
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
                    cmd.CommandText = "spDeleteProductionAllocationForOnScreenEditingProjects";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete Production Allocation
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
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
                            //wbAF.Open(fileFullName);
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the production completed SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> ProductionUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spProductionAllocationDownloadProductionCompletedSKUsForOnScreenEditingProjects";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@ProductionAllocationID", id);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                ProductionUploadedSKUs.Add(sqlReader["UniqueID"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                if (!ProductionUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
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

        #region Production Allocation Details

        #region Download Production Completed Allocation Activities of user for Excel based uploads
        [HttpPost]
        [Route("DownloadProductionCompletedAllocationActivities")]
        public HttpResponseMessage DownloadProductionCompletedAllocationActivities([FromBody] ProductionDownload productionDownload)
        {
            try
            {
                string FileName = "ProductionCompletedActivities_" + productionDownload.ProductionUser + '_' + productionDownload.ProductionAllocationID.ToString() + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationDownloadActivitiesUploads";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionDownload.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", productionDownload.Activities);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionDownload.ProductionUser);

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

        #region Download Production Completed All SKUs of user for Screen Editing Based Projects
        [HttpPost]
        [Route("DownloadProductionCompletedAllSKUsOfUser")]
        public HttpResponseMessage DownloadProductionCompletedAllSKUsOfUser([FromBody] ProductionDownload productionDownload)
        {
            try
            {
                string FileName = "ProductionCompletedAllSKUsOfUser_" + productionDownload.ProductionUser + '_' + productionDownload.ProductionAllocationID.ToString() + ".xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationDownloadProductionCompletedSKUsForOnScreenEditingProjects";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionDownload.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionDownload.ProductionUser);

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

        #region Delete Production Allocation Activities for Excel based uploads
        [HttpPatch]
        [Route("DeleteProductionAllocationActivities")]
        public HttpResponseMessage DeleteProductionAllocationActivities([FromBody] ProductionAllocationDetails productionAllocationDetails)
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
                    cmd.CommandText = "spDeleteProductionAllocationUserActivities";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", productionAllocationDetails.Activities);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionAllocationDetails.ProductionUser);
                    cmd.Parameters.AddWithValue("@UserID", productionAllocationDetails.UserID);

                    //Calling sp to delete Production Allocation Activities
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
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
                            int ProductionUserColumnNo = -1, ActivitiesColNo = -1, UniqueColumnNameColNo = -1;
                            Workbook wbAF = new Workbook();         //Allocated file
                            Aspose.Cells.License l = new Aspose.Cells.License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            //wbAF.Open(fileFullName);
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "production user")
                                    ProductionUserColumnNo = sCol;

                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "activities")
                                    ActivitiesColNo = sCol;

                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (ProductionUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the production uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> ProductionUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spProductionAllocationGetAllProductionUploadedSKUs";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                            cmdPU.Parameters.AddWithValue("@Activities", productionAllocationDetails.Activities);
                            cmdPU.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                ProductionUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                if (ws.Cells[sRow, ProductionUserColumnNo].StringValue.Trim().ToLower() == productionAllocationDetails.ProductionUser.Trim().ToLower() &&
                                    ws.Cells[sRow, ActivitiesColNo].StringValue.Trim().ToLower() == productionAllocationDetails.Activities.Trim().ToLower())
                                {
                                    UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                    if (!ProductionUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
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

        #region Delete Production Allocation SKUs of User for Screen Editing based Projects
        [HttpPatch]
        [Route("DeleteProductionAllocationProductionUsersPendingSKUs")]
        public HttpResponseMessage DeleteProductionAllocationProductionUsersPendingSKUs([FromBody] ProductionAllocationDetails productionAllocationDetails)
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
                    cmd.CommandText = "spDeleteProductionAllocationProductionUsersPendingSKUsForOnScreenEditingProjects";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionAllocationDetails.ProductionUser);
                    cmd.Parameters.AddWithValue("@UserID", productionAllocationDetails.UserID);

                    //Calling sp to delete Production Allocation SKUs of User
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string UniqueColumnName = arrResult[2];
                        AllocatedFileName = arrResult[3];
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
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
                            int ProductionUserColumnNo = -1, UniqueColumnNameColNo = -1;
                            Workbook wbAF = new Workbook();         //Allocated file
                            Aspose.Cells.License l = new Aspose.Cells.License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            //wbAF.Open(fileFullName);
                            wbAF.LoadData(fileFullName);
                            var ws = wbAF.Worksheets[0];

                            for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                            {
                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "production user")
                                    ProductionUserColumnNo = sCol;

                                if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                    UniqueColumnNameColNo = sCol;

                                if (ProductionUserColumnNo >= 0 && UniqueColumnNameColNo >= 0)
                                    break;
                            }

                            #region Fetch all the production completed SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> ProductionUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spProductionAllocationDownloadProductionCompletedSKUsForOnScreenEditingProjects";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                            cmdPU.Parameters.AddWithValue("@ProductionUser", productionAllocationDetails.ProductionUser);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                ProductionUploadedSKUs.Add(sqlReader["UniqueID"].ToString().Trim());
                            conn.Close();
                            #endregion

                            int sRow = 1;
                            int maxRows = ws.Cells.MaxRow;
                            while (sRow <= maxRows)
                            {
                                if (ws.Cells[sRow, ProductionUserColumnNo].StringValue.Trim().ToLower() == productionAllocationDetails.ProductionUser.Trim().ToLower())
                                {
                                    UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                    if (!ProductionUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
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

        #region Change Production User of All Production Pending SKUs for Excel Based Uploads
        [HttpPatch]
        [Route("ChangeUser")]
        public HttpResponseMessage ChangeProductionAllocationUser([FromBody] ProductionAllocationDetails productionAllocationDetails)
        {
            try
            {
                string ChangeToProductionUser = string.Empty;

                if (productionAllocationDetails.ChangeToProductionUser.Contains('-'))
                {
                    string[] arrChangeToProductionUser = productionAllocationDetails.ChangeToProductionUser.Split('-');
                    ChangeToProductionUser = arrChangeToProductionUser[1].Trim();
                }
                else
                    ChangeToProductionUser = productionAllocationDetails.ChangeToProductionUser;

                if (string.IsNullOrEmpty(productionAllocationDetails.Activities))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid activities");

                if (productionAllocationDetails.ProductionUser.Trim().ToLower() == ChangeToProductionUser.Trim().ToLower())
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Production User and Change To Production User cannot be same.");

                if (string.IsNullOrEmpty(productionAllocationDetails.UserID))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid User");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionAllocationChangeProductionUser";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", productionAllocationDetails.Activities);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionAllocationDetails.ProductionUser);
                    cmd.Parameters.AddWithValue("@ChangeToProductionUser", ChangeToProductionUser);
                    cmd.Parameters.AddWithValue("@UserID", productionAllocationDetails.UserID);

                    //Calling sp to change the Production Allocation User
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
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                        string fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        int ProductionUserColumnNo = -1, ActivitiesColNo = -1, UniqueColumnNameColNo = -1;
                        Workbook wbAF = new Workbook();         //Allocated file
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        //wbAF.Open(fileFullName);
                        wbAF.LoadData(fileFullName);
                        var ws = wbAF.Worksheets[0];

                        for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                        {
                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "production user")
                                ProductionUserColumnNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "activities")
                                ActivitiesColNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                UniqueColumnNameColNo = sCol;

                            if (ProductionUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                                break;
                        }

                        if (ProductionUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                        {
                            #region Fetch all the production uploaded SKUs from the allocation
                            string UniqueColumnNameValue = string.Empty;
                            List<string> ProductionUploadedSKUs = new List<string>();
                            System.Data.Common.DbDataReader sqlReader;

                            //Initialize command object
                            SqlCommand cmdPU = new SqlCommand();
                            cmdPU.Connection = conn;
                            cmdPU.CommandType = CommandType.StoredProcedure;
                            cmdPU.CommandText = "spProductionAllocationGetAllProductionUploadedSKUs";

                            //Add parameters with values
                            cmdPU.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationDetails.ProductionAllocationID);
                            cmdPU.Parameters.AddWithValue("@Activities", productionAllocationDetails.Activities);
                            cmdPU.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);

                            //Calling sp to get list of SKUs
                            conn.Open();
                            cmdPU.CommandTimeout = 0;
                            sqlReader = (System.Data.Common.DbDataReader)cmdPU.ExecuteReader();
                            while (sqlReader.Read())
                                ProductionUploadedSKUs.Add(sqlReader["UniqueColumnName"].ToString().Trim());
                            conn.Close();
                            #endregion

                            for (int sRow = 1; sRow <= ws.Cells.MaxRow; sRow++)
                            {
                                if (ws.Cells[sRow, ProductionUserColumnNo].StringValue.Trim().ToLower() == productionAllocationDetails.ProductionUser.Trim().ToLower() &&
                                        ws.Cells[sRow, ActivitiesColNo].StringValue.Trim().ToLower() == productionAllocationDetails.Activities.Trim().ToLower())
                                {
                                    UniqueColumnNameValue = ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim();
                                    if (!ProductionUploadedSKUs.Contains(UniqueColumnNameValue, StringComparer.OrdinalIgnoreCase))
                                        ws.Cells[sRow, ProductionUserColumnNo].PutValue(ChangeToProductionUser);
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

        #region Production Pending SKUs By Production User for Excel Based Uploads
        [HttpPost]
        [Route("ReadProductionPendingSKUsByProductionUser")]
        public IHttpActionResult ReadProductionPendingSKUsByProductionUser([FromBody] ProductionDownload productionDownload)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;
                DataTable dtProdPendingSKUs = new DataTable();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationGetPendingSKUsByProdUser";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionDownload.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@Activities", productionDownload.Activities);
                    cmd.Parameters.AddWithValue("@ProductionUser", productionDownload.ProductionUser);

                    //Calling sp to get uploaded data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        dtProdPendingSKUs.Load(sqlReader);
                    }
                    conn.Close();

                    //return datatable to the request
                    return Ok(dtProdPendingSKUs);
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

        #region Change SKU Production User for Excel Based Uploads
        [HttpPatch]
        [Route("ChangeSKUProductionUser")]
        public HttpResponseMessage ChangeSKUProductionUser([FromBody] ProductionAllocationChangeSKUProdUser productionAllocationChangeSKUProdUser)
        {
            try
            {
                if (productionAllocationChangeSKUProdUser.ProductionAllocationID == 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid production allocation id");

                if (string.IsNullOrEmpty(productionAllocationChangeSKUProdUser.UniqueColumnValue))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid unique column value");

                if (string.IsNullOrEmpty(productionAllocationChangeSKUProdUser.Activities))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid activities");

                if (string.IsNullOrEmpty(productionAllocationChangeSKUProdUser.ChangeToProductionUser) || productionAllocationChangeSKUProdUser.ChangeToProductionUser.Length != 3)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid change to production user");

                if (string.IsNullOrEmpty(productionAllocationChangeSKUProdUser.UserID) || productionAllocationChangeSKUProdUser.UserID.Length != 3)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid login User");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProductionAllocationChangeSKUProdUser";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", productionAllocationChangeSKUProdUser.ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@UniqueColumnValue", productionAllocationChangeSKUProdUser.UniqueColumnValue);
                    cmd.Parameters.AddWithValue("@Activities", productionAllocationChangeSKUProdUser.Activities);
                    cmd.Parameters.AddWithValue("@ChangeToProductionUser", productionAllocationChangeSKUProdUser.ChangeToProductionUser);
                    cmd.Parameters.AddWithValue("@UserID", productionAllocationChangeSKUProdUser.UserID);

                    //Calling sp to change the Production Allocated User
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
                        string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                        string fileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                        int ProductionUserColumnNo = -1, ActivitiesColNo = -1, UniqueColumnNameColNo = -1;
                        Workbook wbAF = new Workbook();         //Allocated file
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        //wbAF.Open(fileFullName);
                        wbAF.LoadData(fileFullName);
                        var ws = wbAF.Worksheets[0];

                        for (int sCol = 0; sCol <= ws.Cells.MaxColumn; sCol++)
                        {
                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "production user")
                                ProductionUserColumnNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == "activities")
                                ActivitiesColNo = sCol;

                            if (ws.Cells[0, sCol].StringValue.Trim().ToLower() == UniqueColumnName.ToLower())
                                UniqueColumnNameColNo = sCol;

                            if (ProductionUserColumnNo >= 0 && ActivitiesColNo >= 0 && UniqueColumnNameColNo >= 0)
                                break;
                        }

                        if (UniqueColumnNameColNo >= 0 && ActivitiesColNo >= 0 && ProductionUserColumnNo >= 0)
                        {
                            for (int sRow = 1; sRow <= ws.Cells.MaxRow; sRow++)
                            {
                                if (ws.Cells[sRow, UniqueColumnNameColNo].StringValue.Trim().ToLower() == productionAllocationChangeSKUProdUser.UniqueColumnValue.Trim().ToLower() &&
                                        ws.Cells[sRow, ActivitiesColNo].StringValue.Trim().ToLower() == productionAllocationChangeSKUProdUser.Activities.Trim().ToLower())
                                {
                                    ws.Cells[sRow, ProductionUserColumnNo].PutValue(productionAllocationChangeSKUProdUser.ChangeToProductionUser);
                                    break;
                                }
                            }
                            wbAF.Save(fileFullName);
                        }
                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].ToLower());
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

        #region Download All Allocated SKUs of Project for both Excel Based Uploads and On Screen Editing Projects
        [HttpGet]
        [Route("DownloadAllProductionAllocatedSKUs/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public HttpResponseMessage DownloadAllProductionAllocatedSKUs(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string tempPath = System.IO.Path.GetTempPath();
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                string ProductionAllocatedFileName = string.Empty, ProductionAllocatedFileFullName = string.Empty, AllocatedFileColumnHeader = string.Empty;
                bool IsFirstFile = true, IsAllocatedFileColumnFoundInDownloadFile = false;
                string DownloadFileName = "ProductionAllocation_" + CustomerCode + '_' + ProjectCode;
                if (!string.IsNullOrEmpty(BatchNo))
                    DownloadFileName = DownloadFileName + '_' + BatchNo;

                DownloadFileName = DownloadFileName + ".xlsx";

                string DownloadFileFullName = tempPath + DownloadFileName;
                int DownloadFileLastRowNo = 0, DownloadFileLastColumnNo = -1;

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                DirectoryInfo dirProductionAllocationUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/"));
                foreach (FileInfo file in dirProductionAllocationUploads.GetFiles())
                {
                    if (file.Name.ToLower().StartsWith("productionallocation_" + CustomerCode.ToLower() + '_' + ProjectCode.ToLower()))
                    {
                        if (!string.IsNullOrEmpty(BatchNo))
                        {
                            if (file.Name.ToLower().StartsWith("productionallocation_" + CustomerCode.ToLower() + '_' + ProjectCode.ToLower() + '_' + BatchNo.ToLower()))
                                ProductionAllocatedFileName = file.Name;
                            else
                                ProductionAllocatedFileName = string.Empty;
                        }
                        else
                            ProductionAllocatedFileName = file.Name;

                        ProductionAllocatedFileFullName = Path.Combine(fileUploadedPath, ProductionAllocatedFileName);

                        if (!File.Exists(ProductionAllocatedFileFullName))
                            continue;

                        if (IsFirstFile)
                        {
                            if (File.Exists(DownloadFileFullName))
                                File.Delete(DownloadFileFullName);

                            file.CopyTo(DownloadFileFullName);

                            IsFirstFile = false;
                        }
                        else
                        {
                            //second file onwards read here
                            Workbook wbAF = new Workbook();
                            License l = new License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            //wbAF.Open(ProductionAllocatedFileFullName);
                            wbAF.LoadData(ProductionAllocatedFileFullName);
                            Worksheet wsAF = wbAF.Worksheets[0];

                            Workbook wbDF = new Workbook();
                            License l1 = new License();
                            l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            //wbDF.Open(DownloadFileFullName);
                            wbDF.LoadData(DownloadFileFullName);
                            Worksheet wsDF = wbDF.Worksheets[0];

                            DownloadFileLastRowNo = wsDF.Cells.MaxRow;
                            DownloadFileLastColumnNo = wsDF.Cells.MaxColumn;
                            for (int AFRow = 1; AFRow <= wsAF.Cells.MaxRow; AFRow++)
                            {
                                DownloadFileLastRowNo++;
                                for (int AFCol = 0; AFCol <= wsAF.Cells.MaxColumn; AFCol++)
                                {
                                    AllocatedFileColumnHeader = wsAF.Cells[0, AFCol].StringValue.Trim();

                                    IsAllocatedFileColumnFoundInDownloadFile = false;
                                    for (int DFCol = 0; DFCol <= wsDF.Cells.MaxColumn; DFCol++)
                                    {
                                        if (wsDF.Cells[0, DFCol].StringValue.ToLower().Trim() == AllocatedFileColumnHeader.ToLower().Trim())
                                        {
                                            wsDF.Cells[DownloadFileLastRowNo, DFCol].PutValue(wsAF.Cells[AFRow, AFCol].StringValue.Trim());
                                            IsAllocatedFileColumnFoundInDownloadFile = true;
                                            break;
                                        }
                                    }

                                    if (!IsAllocatedFileColumnFoundInDownloadFile)
                                    {
                                        DownloadFileLastColumnNo++;
                                        wsDF.Cells[0, DownloadFileLastColumnNo].PutValue(AllocatedFileColumnHeader);
                                        wsDF.Cells[DownloadFileLastRowNo, DownloadFileLastColumnNo].PutValue(wsAF.Cells[AFRow, AFCol].StringValue.Trim());
                                    }
                                }
                            }

                            wbDF.Save(DownloadFileFullName);
                        }
                    }
                }

                //Read the file into a Byte Array.
                byte[] bytes = File.ReadAllBytes(DownloadFileFullName);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = DownloadFileName;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(DownloadFileName));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Download All Production Pending SKUs for Excel Based Uploads
        [HttpGet]
        [Route("DownloadAllProductionPendingSKUs/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public HttpResponseMessage DownloadAllProductionPendingSKUs(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string tempPath = System.IO.Path.GetTempPath();
                string fileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/");
                string ProductionAllocatedFileFullName = string.Empty;
                string DownloadFileName = "ProductionPendingSKUs_" + CustomerCode + '_' + ProjectCode;
                if (!string.IsNullOrEmpty(BatchNo))
                    DownloadFileName = DownloadFileName + '_' + BatchNo;

                DownloadFileName = DownloadFileName + ".xlsx";

                string DownloadFileFullName = tempPath + DownloadFileName;
                int DownloadFileLastRowNo = 0, DownloadFileLastColumnNo = -1;

                string UniqueColumnName = string.Empty;
                string AllocatedFileName = string.Empty, AllocatedFileNameWithoutExtension = string.Empty, AllocatedFileColumnHeader = string.Empty;
                bool IsToSearchInAllocatedFile = false, IsAllocatedFileColumnFoundInDownloadFile = false, IsToWriteFirstTime = true;
                long AllocatedFileProductionAllocationID = 0;

                int UniqueColumnIndex = -1, ActivitiesColumnIndex = -1;

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                #region Fetch Production Pending SKUs
                List<ProductionPendingSKU> productionPendingSKUsList = new List<ProductionPendingSKU>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationGetAllProductionPendingSKUs";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    //Calling sp to get list of existing Production Allocations
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProductionPendingSKU productionPendingSKUs = new ProductionPendingSKU();

                        productionPendingSKUs.ProductionAllocationID = Convert.ToInt64(sqlReader["ProductionAllocationID"]);
                        productionPendingSKUs.UniqueColumnName = sqlReader["UniqueColumnName"].ToString();
                        productionPendingSKUs.UniqueColumnNameValues = sqlReader["UniqueColumnNameValues"].ToString();
                        productionPendingSKUs.Activity = sqlReader["Activity"].ToString();

                        productionPendingSKUsList.Add(productionPendingSKUs);
                    }
                    conn.Close();
                }
                #endregion

                #region Open Download File
                Workbook wbDF = new Workbook();
                License l1 = new License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wbDF.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wbDF.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                Worksheet wsDF = wbDF.Worksheets[0];
                #endregion

                #region Search Production Pending SKUs in Allocated File and Write to Download File
                if (productionPendingSKUsList.Count() > 0)
                {
                    UniqueColumnName = productionPendingSKUsList.FirstOrDefault().UniqueColumnName.Trim();

                    DirectoryInfo dirProductionAllocationUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProductionAllocation/"));
                    foreach (FileInfo file in dirProductionAllocationUploads.GetFiles())
                    {
                        AllocatedFileName = file.Name.Trim();
                        AllocatedFileNameWithoutExtension = AllocatedFileName.Substring(0, file.Name.IndexOf('.'));

                        #region Check Allocated File Name
                        string[] ArrayOfAllocatedFileName = AllocatedFileNameWithoutExtension.Split('_');
                        if (ArrayOfAllocatedFileName[0].Trim().ToLower() == "productionallocation" &&
                           ArrayOfAllocatedFileName[1].Trim().ToLower() == CustomerCode.Trim().ToLower() &&
                           ArrayOfAllocatedFileName[2].Trim().ToLower() == ProjectCode.Trim().ToLower())
                        {
                            if (!string.IsNullOrEmpty(BatchNo))
                            {
                                if (ArrayOfAllocatedFileName[3].Trim().ToLower() == BatchNo.Trim().ToLower())
                                    AllocatedFileProductionAllocationID = Convert.ToInt64(ArrayOfAllocatedFileName[4].Trim());
                            }
                            else
                                AllocatedFileProductionAllocationID = Convert.ToInt64(ArrayOfAllocatedFileName[3].Trim());

                            if (productionPendingSKUsList.Count(p => p.ProductionAllocationID == AllocatedFileProductionAllocationID) > 0)
                                IsToSearchInAllocatedFile = true;
                            else
                                IsToSearchInAllocatedFile = false;
                        }
                        else
                            IsToSearchInAllocatedFile = false;
                        #endregion

                        if (IsToSearchInAllocatedFile)
                        {
                            ProductionAllocatedFileFullName = Path.Combine(fileUploadedPath, AllocatedFileName);

                            #region Open Allocated File
                            Workbook wbAF = new Workbook();
                            License l = new License();
                            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                            //wbAF.Open(ProductionAllocatedFileFullName);
                            wbAF.LoadData(ProductionAllocatedFileFullName);
                            Worksheet wsAF = wbAF.Worksheets[0];
                            #endregion

                            #region Search Unique Column Name and Activities columns in Allocated File and Get the column index
                            UniqueColumnIndex = -1; ActivitiesColumnIndex = -1;
                            for (int AFcol = 0; AFcol <= wsAF.Cells.MaxColumn; AFcol++)
                            {
                                if (wsAF.Cells[0, AFcol].StringValue.Trim().ToLower() == UniqueColumnName.Trim().ToLower())
                                    UniqueColumnIndex = AFcol;

                                if (wsAF.Cells[0, AFcol].StringValue.Trim().ToLower() == "activities")
                                    ActivitiesColumnIndex = AFcol;

                                if (UniqueColumnIndex >= 0 && ActivitiesColumnIndex >= 0)
                                    break;
                            }
                            #endregion

                            #region Check Unique Column Name and Activity exists in Allocated File and write that row to download file
                            DownloadFileLastRowNo = wsDF.Cells.MaxRow;
                            if (IsToWriteFirstTime)
                                DownloadFileLastColumnNo = -1;
                            else
                                DownloadFileLastColumnNo = wsDF.Cells.MaxColumn;
                            for (int AFRow = 1; AFRow <= wsAF.Cells.MaxRow; AFRow++)
                            {
                                string[] ArrayOfAllocatedActivities = wsAF.Cells[AFRow, ActivitiesColumnIndex].StringValue.Trim().Split('|').Select(e => e.Trim()).ToArray();
                                if (productionPendingSKUsList.Count(p => p.UniqueColumnNameValues.Trim().ToLower() ==
                                    wsAF.Cells[AFRow, UniqueColumnIndex].StringValue.Trim().ToLower() &&
                                    ArrayOfAllocatedActivities.Contains(p.Activity.Trim(), StringComparer.OrdinalIgnoreCase)) > 0)
                                {
                                    DownloadFileLastRowNo++;
                                    for (int AFCol = 0; AFCol <= wsAF.Cells.MaxColumn; AFCol++)
                                    {
                                        AllocatedFileColumnHeader = wsAF.Cells[0, AFCol].StringValue.Trim();
                                        IsAllocatedFileColumnFoundInDownloadFile = false;
                                        for (int DFCol = 0; DFCol <= wsDF.Cells.MaxColumn; DFCol++)
                                        {
                                            if (wsDF.Cells[0, DFCol].StringValue.ToLower().Trim() == AllocatedFileColumnHeader.ToLower().Trim())
                                            {
                                                wsDF.Cells[DownloadFileLastRowNo, DFCol].PutValue(wsAF.Cells[AFRow, AFCol].StringValue.Trim());
                                                IsAllocatedFileColumnFoundInDownloadFile = true;
                                                break;
                                            }
                                        }

                                        if (!IsAllocatedFileColumnFoundInDownloadFile)
                                        {
                                            DownloadFileLastColumnNo++;
                                            wsDF.Cells[0, DownloadFileLastColumnNo].PutValue(AllocatedFileColumnHeader);
                                            wsDF.Cells[DownloadFileLastRowNo, DownloadFileLastColumnNo].PutValue(wsAF.Cells[AFRow, AFCol].StringValue.Trim());
                                        }
                                    }
                                    IsToWriteFirstTime = false;
                                }
                            }
                            #endregion
                        }
                        wbDF.Save(DownloadFileFullName);
                    }

                    #region Download File
                    //Read the file into a Byte Array.
                    byte[] bytes = File.ReadAllBytes(DownloadFileFullName);

                    //Set the Response Content.
                    response.Content = new ByteArrayContent(bytes);

                    //Set the Response Content Length.
                    response.Content.Headers.ContentLength = bytes.LongLength;

                    //Set the Content Disposition Header Value and FileName.
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                    response.Content.Headers.ContentDisposition.FileName = DownloadFileName;

                    //Set the File Content Type.
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(DownloadFileName));
                    #endregion

                    return response;
                }
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No data found");

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

        #region Download Production Pending SKUs for On Screen Editing Projects
        [HttpGet]
        [Route("DownloadProductionPendingSKUs")]
        public HttpResponseMessage DownloadProductionPendingSKUs(string CustomerCode, string ProjectCode, string BatchNo = "", long ProductionAllocationID = 0, string ProductionUser = "")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string FileName = "Production Pending SKUs.xlsx";

                System.Data.Common.DbDataReader sqlReader;
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var ws = wb.Worksheets[0];
                int row = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProductionAllocationProductionPendingSKUsForOnScreenEditingProjects";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionAllocationID", ProductionAllocationID);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);

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

        #endregion

        #region Download Help Document file
        [HttpGet]
        [Route("DownloadHelpDocument")]
        public HttpResponseMessage DownloadHelpDocument()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileName = "GOP Help Document.docx";

                //Set the Help File Path
                string filePath = HttpContext.Current.Server.MapPath("~/HelpDocs/") + FileName;

                //Check whether File exists.
                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.NotFound, "File not found");

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(filePath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileName));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region On-Screen Production Allocation
        #region Read Unique Column Names from Customer Input File
        [HttpGet]
        [Route("ReadUniqueColumnNamesFromCIF")]
        public IHttpActionResult ReadUniqueColumnNamesFromCIF(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            string FileName, InputFilepath;

            //Swagger fix
            if (BatchNo == "{BatchNo}")
                BatchNo = "";

            if (string.IsNullOrEmpty(BatchNo))
            {
                FileName = CustomerCode + "_" + ProjectCode + "_CustomerInputFile.xlsx";
                InputFilepath = HttpContext.Current.Server.MapPath(@"\Uploads\GPMT\CustomerInputFile\" + FileName);
            }
            else
            {
                FileName = CustomerCode + "_" + ProjectCode + "_" + BatchNo + "_CustomerInputFile.xlsx";
                InputFilepath = HttpContext.Current.Server.MapPath(@"\Uploads\GPMT\CustomerInputFile\Batch\" + FileName);
            }

            //Create a list to hold the list of Unique Column Names
            List<UniqueColumn> UniqueColumnNamesList = new List<UniqueColumn>();

            #region Read unique Column Names from file
            Workbook wbIF = new Workbook();                       //input file work book
            Aspose.Cells.License l = new Aspose.Cells.License();
            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
            wbIF.LoadData(InputFilepath);
            Worksheet wsIF = wbIF.Worksheets[0];

            for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
            {
                UniqueColumn uniqueColumn = new UniqueColumn();
                uniqueColumn.UniqueColumnName = wsIF.Cells[0, col].StringValue.Trim();
                if (!UniqueColumnNamesList.Any(u => u.UniqueColumnName.Trim().ToLower() == uniqueColumn.UniqueColumnName.ToLower()))
                    UniqueColumnNamesList.Add(uniqueColumn);
            }

            return Ok(UniqueColumnNamesList);
            #endregion
        }
        #endregion

        #region Fetch Pending To Allocate SKUs from Customer Input File
        [HttpGet]
        [Route("ReadPendingToAllocateSKUsFromCIF")]
        public IHttpActionResult ReadPendingToAllocateSKUsFromCIF(string CustomerCode, string ProjectCode, string BatchNo = "", string SearchOn = "", string SearchText = "")
        {
            try
            {
                DataTable dt = null;
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOSPAFetchSKUs";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@SearchOn", SearchOn);
                    cmd.Parameters.AddWithValue("@SearchText", SearchText);
                    cmd.Parameters.AddWithValue("@WhichSKUs", "P");

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

        #region Validate And Allocate Pending SKUs
        [HttpPost]
        [Route("ValidateAndAllocatePendingSKUs")]
        public HttpResponseMessage ValidateAndAllocatePendingSKUs([FromBody] OSPASKUs model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Pending SKUs model data");

                #region CIF SKU IDs
                var serializer = new XmlSerializer(typeof(List<CIFSKUID>),
                                       new XmlRootAttribute("root"));
                var stream = new StringWriter();
                serializer.Serialize(stream, model.CIFIDs);
                StringReader transactionXml = new StringReader(stream.ToString());
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spOSPAValidateAndAllocatePendingSKUs";

                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.Add(new SqlParameter("@CIFIDs", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@AllocateToUser", model.AllocateToUser);
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "created")
                        return Request.CreateResponse(HttpStatusCode.OK, "SKUs allocated successfully to user");
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

        #region Read Production Allocated SKUs of Project
        [HttpGet]
        [Route("ReadProductionAllocatedSKUsOfProject")]
        public IHttpActionResult ReadProductionAllocatedSKUsOfProject(string CustomerCode, string ProjectCode, string ProductionUser = "", string BatchNo = "", string WhichSKUs = "A")
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                List<OSPAAllocatedSKUs> allocatedSKUsList = new List<OSPAAllocatedSKUs>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOSPAFetchSKUs";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@ProductionUser", ProductionUser);
                    cmd.Parameters.AddWithValue("@WhichSKUs", WhichSKUs);

                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        OSPAAllocatedSKUs ospaAllocatedSKU = new OSPAAllocatedSKUs();
                        ospaAllocatedSKU.ID = Convert.ToInt32(sqlReader["ID"].ToString());
                        ospaAllocatedSKU.ShortDescription = sqlReader["ShortDescription"].ToString();
                        ospaAllocatedSKU.LongDescription = sqlReader["LongDescription"].ToString();
                        ospaAllocatedSKU.UOM = sqlReader["UOM"].ToString();
                        ospaAllocatedSKU.MFRName = sqlReader["MFRName"].ToString();
                        ospaAllocatedSKU.MFRPN = sqlReader["MFRPN"].ToString();
                        ospaAllocatedSKU.VendorName = sqlReader["VendorName"].ToString();
                        ospaAllocatedSKU.VendorPN = sqlReader["VendorPN"].ToString();
                        ospaAllocatedSKU.Status = sqlReader["Status"].ToString();
                        ospaAllocatedSKU.ProductionUser = sqlReader["ProductionUser"].ToString();
                        allocatedSKUsList.Add(ospaAllocatedSKU);
                    }
                    conn.Close();

                    return Ok(allocatedSKUsList);
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

        #region Validate and Move Allocated SKUs To Pending
        [HttpPost]
        [Route("ValidateAndMoveAllocatedSKUsToPending")]
        public HttpResponseMessage ValidateAndMoveAllocatedSKUsToPending([FromBody] OSPASKUs model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Allocated SKUs model data");

                #region CIF SKU IDs
                var serializer = new XmlSerializer(typeof(List<CIFSKUID>),
                                       new XmlRootAttribute("root"));
                var stream = new StringWriter();
                serializer.Serialize(stream, model.CIFIDs);
                StringReader transactionXml = new StringReader(stream.ToString());
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spOSPAMoveAllocatedSKUsToPending";

                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.Add(new SqlParameter("@CIFIDs", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, "Allocated SKUs moved to pending successfully");
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

        #region Read Project Allocated User Names
        [HttpGet]
        [Route("ReadProjectAllocatedUserNames")]
        public HttpResponseMessage ReadProjectAllocatedUserNames(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                DataTable dataTable = new DataTable();
                List<string> userNames = new List<string>();

                using (SqlConnection con = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    using (SqlCommand cmd = new SqlCommand("spOSPAFetchUserNames", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }

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

        #region Validate and Re-Allocate SKUs
        [HttpPost]
        [Route("ValidateAndReAllocateSKUs")]
        public HttpResponseMessage ValidateAndReAllocateSKUs([FromBody] OSPASKUs model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid allocated SKUs model data");

                #region CIF SKU IDs
                var serializer = new XmlSerializer(typeof(List<CIFSKUID>),
                                       new XmlRootAttribute("root"));
                var stream = new StringWriter();
                serializer.Serialize(stream, model.CIFIDs);
                StringReader transactionXml = new StringReader(stream.ToString());
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spOSPAValidateAndReAllocateSKUs";

                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                    cmd.Parameters.Add(new SqlParameter("@CIFIDs", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@ReAllocateToUser", model.AllocateToUser);
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);

                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, "SKUs re-allocated successfully to user");
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
    }
}
