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
using Aspose.Cells;
using System.Data.OleDb;
using System.Data.Common;
using GOPWebAPI.Helpers;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/projectallocation")]
    public class ProjectAllocationController : ApiController
    {
        #region Download Project Allocation Template File
        [HttpGet]
        [Route("downloadallocationtemplate")]
        public HttpResponseMessage DownloadAllocationTemplate()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileName = "ProjectAllocationTemplate.xlsx";

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/Templates/" + FileName);

                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Project Allocation Template file not found");

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

        #region Upload input file
        // call api/project/uploadfile URL to Upload input file to temp folder
        #endregion

        #region Validate Input File
        [HttpGet]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile()          //string FileName
        {
            try
            {
                string FileName = "2d65d18e-35c8-4865-98a2-026b8908c50f.xlsx";            //temporary
                string CustomerCode="DLG";                                                //temporary
                string ProjectCode = "019";                                               //temporary
                string BatchNo = "0002";                                                  //temporary
                string TemplateFilePath = HttpContext.Current.Server.MapPath(@"\Templates\ProjectAllocationTemplate.xlsx");
                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (!File.Exists(InputFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file not found");

                Workbook TFwb = new Workbook();                       //template file work book
                Aspose.Cells.License tl = new Aspose.Cells.License();
                tl.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                TFwb.Open(TemplateFilePath);
                Worksheet TFws = TFwb.Worksheets[0];
                
                Workbook IFwb = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilepath);
                Worksheet IFws = IFwb.Worksheets[0];

                #region Checking input file for duplicate columns if any
                for(int mcol=0;mcol<IFws.Cells.MaxColumn-1;mcol++)              //main column
                {
                    for(int ncol=mcol+1;ncol<=IFws.Cells.MaxColumn;ncol++)      //navigation column
                    {
                        if(IFws.Cells[0,mcol].StringValue.Trim().ToUpper()==IFws.Cells[0,ncol].StringValue.Trim().ToUpper())
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file has duplicate columns. column name:" + IFws.Cells[0, mcol].StringValue.Trim());
                    }
                }
                #endregion

                #region Checking input file to have columns as per template
                for(int col=0;col<=10;col++)                //since first 11 columns must be as per template, it is checked against template
                {
                    if(IFws.Cells[0,col].StringValue.Trim().ToUpper()!=TFws.Cells[0,col].StringValue.Trim().ToUpper())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file columns does not match to template file columns. column name:" + IFws.Cells[0, col].StringValue.Trim());
                }
                #endregion

                #region Checking input file has data rows
                if (IFws.Cells.MaxRow <= 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file has no data rows.");
                #endregion

                #region Checking all columns have Production Employee Code, QC Employee Code, QA Employee Code if exists
                for(int row=1;row<=IFws.Cells.MaxRow;row++)
                {
                    //if all columns Prod. emp. code, QC emp. code, QA emp. code are empty then continue
                    if (string.IsNullOrEmpty(IFws.Cells[row, 3].StringValue.Trim()) &&
                        string.IsNullOrEmpty(IFws.Cells[row, 4].StringValue.Trim()) &&
                        string.IsNullOrEmpty(IFws.Cells[row, 5].StringValue.Trim()))
                        continue;
                    
                    //if Production emp. is allocated, QC and QA emp. must be allocated
                    if(!string.IsNullOrEmpty(IFws.Cells[row, 3].StringValue.Trim()) &&
                       (string.IsNullOrEmpty(IFws.Cells[row, 4].StringValue.Trim()) ||
                        string.IsNullOrEmpty(IFws.Cells[row, 5].StringValue.Trim())))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "If Production emp. is allocated, QC and QA emp. must be allocated");

                    //if QC emp. is allocated, Production and QA emp. must be allocated
                    if (!string.IsNullOrEmpty(IFws.Cells[row, 4].StringValue.Trim()) &&
                       (string.IsNullOrEmpty(IFws.Cells[row, 3].StringValue.Trim()) ||
                        string.IsNullOrEmpty(IFws.Cells[row, 5].StringValue.Trim())))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "If QC emp. is allocated, Production and QA emp. must be allocated");

                    //if QA emp. is allocated, QC and Production emp. must be allocated
                    if (!string.IsNullOrEmpty(IFws.Cells[row, 5].StringValue.Trim()) &&
                       (string.IsNullOrEmpty(IFws.Cells[row, 3].StringValue.Trim()) ||
                        string.IsNullOrEmpty(IFws.Cells[row, 4].StringValue.Trim())))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "If QA emp. is allocated, QC and Production emp. must be allocated");

                }
                #endregion

                #region Checking Customer Code, Project Code, Batch No.
                for(int row=1;row<=IFws.Cells.MaxRow;row++)
                {
                    if(IFws.Cells[row,0].StringValue.Trim().ToLower()!=CustomerCode.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "All rows from input file should have customer code:" + CustomerCode);

                    if (IFws.Cells[row, 1].StringValue.Trim().ToLower() != ProjectCode.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "All rows from input file should have project code:" + ProjectCode);

                    if (IFws.Cells[row, 2].StringValue.Trim().ToLower() != BatchNo.Trim().ToLower())
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Batch no. of all rows from input file not matching to input batch no.");
                }
                #endregion

                #region Checking Input File Has Valid Activities
                List<string> InputFileActivities = new List<string>();
                for(int row=1;row<=IFws.Cells.MaxRow;row++)
                {
                    if(!string.IsNullOrEmpty(IFws.Cells[row,10].StringValue.Trim()))
                        InputFileActivities.Add(IFws.Cells[row, 10].StringValue.Trim());
                }

                InputFileActivities = InputFileActivities.Distinct<string>().ToList();
                
                var serializer = new XmlSerializer(typeof(List<string>),
                                       new XmlRootAttribute("root"));
                var stream = new StringWriter();
                serializer.Serialize(stream, InputFileActivities);

                StringReader transactionXml = new StringReader(stream.ToString());
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);
                SqlXml sqlXml = new SqlXml(xmlReader);

                string ConnectionString = DBConnInfo.ConnectionString();
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spValidateProjectActivities";

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.Add(new SqlParameter("@Activities", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if(Result.ToLower()!="success")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
                #endregion

                var message = Request.CreateResponse(HttpStatusCode.OK, "validation successful");
                return message;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Employeewise Summary
        [HttpGet]
        public IHttpActionResult ReadEmployeewiseSummary()              //string FileName
        {
            string FileName = "2d65d18e-35c8-4865-98a2-026b8908c50f.xlsx";
            List<string> EmployeeCodes=new List<string>();
            List<EmployeeAllocation> employeesAllocation = new List<EmployeeAllocation>();

            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
            Workbook IFwb = new Workbook();                       //input file work book
            Aspose.Cells.License l = new Aspose.Cells.License();
            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
            IFwb.Open(InputFilepath);
            Worksheet IFws = IFwb.Worksheets[0];
            
            for(int row=1;row<=IFws.Cells.MaxRow;row++)
            {
                //Adding Production Employees
                if (!string.IsNullOrEmpty(IFws.Cells[row, 3].StringValue.Trim()))
                {
                    if (!EmployeeCodes.Contains(IFws.Cells[row, 3].StringValue.Trim()))
                        EmployeeCodes.Add(IFws.Cells[row, 3].StringValue.Trim().ToUpper());
                }

                //Adding QC Employees
                if (!string.IsNullOrEmpty(IFws.Cells[row, 4].StringValue.Trim()))
                {
                    if (!EmployeeCodes.Contains(IFws.Cells[row, 4].StringValue.Trim()))
                        EmployeeCodes.Add(IFws.Cells[row, 4].StringValue.Trim().ToUpper());
                }

                //Adding QA Employees
                if (!string.IsNullOrEmpty(IFws.Cells[row, 5].StringValue.Trim()))
                {
                    if (!EmployeeCodes.Contains(IFws.Cells[row, 5].StringValue.Trim()))
                        EmployeeCodes.Add(IFws.Cells[row, 5].StringValue.Trim().ToUpper());
                }
            }

            foreach(string empcode in EmployeeCodes)
                employeesAllocation.Add(new EmployeeAllocation() { EmployeeCode = empcode, ProductionAllocated = 0, QCAllocated = 0, QAAllocated = 0 });
            
            //Update Production Count
            for(int row=1;row<=IFws.Cells.MaxRow;row++)
            {
                var foundRow = employeesAllocation.FirstOrDefault(e => e.EmployeeCode.ToLower() == IFws.Cells[row,3].StringValue.Trim().ToLower());
                foundRow.ProductionAllocated = foundRow.ProductionAllocated + 1;
            }

            //Update QC Count
            for (int row = 1; row <= IFws.Cells.MaxRow; row++)
            {
                var foundRow = employeesAllocation.FirstOrDefault(e => e.EmployeeCode.ToLower() == IFws.Cells[row, 4].StringValue.Trim().ToLower());
                foundRow.QCAllocated = foundRow.QCAllocated + 1;
            }

            //Update QA Count
            for (int row = 1; row <= IFws.Cells.MaxRow; row++)
            {
                var foundRow = employeesAllocation.FirstOrDefault(e => e.EmployeeCode.ToLower() == IFws.Cells[row, 5].StringValue.Trim().ToLower());
                foundRow.QAAllocated = foundRow.QAAllocated + 1;
            }

            return Ok(employeesAllocation);
        }
        #endregion

        #region Read Unique Column Names
        [HttpGet]
        [Route("ReadUniqueColumnNames")]
        public IHttpActionResult ReadUniqueColumnNames(string FileName)
        {
            try
            {
                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
                List<string> UniqueColumnNamesList = new List<string>();

                Workbook IFwb = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilepath);
                Worksheet IFws = IFwb.Worksheets[0];

                for (int col = 0; col <= IFws.Cells.MaxColumn; col++)
                    UniqueColumnNamesList.Add(IFws.Cells[0, col].StringValue.Trim());

                return Ok(UniqueColumnNamesList);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Create Allocation
        [HttpPost]
        [Route]
        public HttpResponseMessage PostProjectAllocation()          //[FromBody] ProjectAllocation projectAllocation
        {
            try
            {
                string FileName = "2d65d18e-35c8-4865-98a2-026b8908c50f.xlsx";              //temporary
                string CustomerCode = "DLG";                                                //temporary
                string ProjectCode = "019";                                                 //temporary
                string BatchNo = "0002";                                                    //temporary
                string UniqueColumnName = "Material Code";                                  //temporary
                string UserID = "vic";                                                      //temporary

                //if(!ModelState.IsValid)
                //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Allocation Data");
                
                //string FileName = projectAllocation.FileName;      
                //string CustomerCode = projectAllocation.CustomerCode;
                //string ProjectCode = projectAllocation.ProjectCode;
                //string BatchNo = projectAllocation.BatchNo;
                //string UniqueColumnName = projectAllocation.UniqueColumnName;
                //string UserID = projectAllocation.UserID;

                string sqlTableName = "ProjectAllocation_";
                string ConnectionString = DBConnInfo.ConnectionString();
                string ErrorMessage = string.Empty;
                long ProjectAllocationID = 0;
                string InputFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                Workbook IFwb = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilePath);
                Worksheet IFws = IFwb.Worksheets[0];

                #region Check all values from Unique Column are unique
                int UniqueColumnNo = -1;
                for(int col=0;col<=IFws.Cells.MaxColumn;col++)
                {
                    if(IFws.Cells[0,col].StringValue.Trim().ToLower()==UniqueColumnName.Trim().ToLower())
                    {
                        UniqueColumnNo = col;
                        break;
                    }
                }

                if(UniqueColumnNo<0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Unique column name not found in input file");

                List<string> UniqueColumnsAllValuesList = new List<string>();
                for(int row=1;row<=IFws.Cells.MaxRow;row++)
                {
                    if(string.IsNullOrEmpty(IFws.Cells[row,UniqueColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Unique column name all values are mandatory");

                    UniqueColumnsAllValuesList.Add(IFws.Cells[row, UniqueColumnNo].StringValue.Trim());
                }

                List<string> UniqueColumnsDistinctValuesList = UniqueColumnsAllValuesList.Distinct<string>().ToList();
                if (UniqueColumnsDistinctValuesList.Count != UniqueColumnsAllValuesList.Count)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Duplicate values found in unique column");
                #endregion
                
                #region Add entry in Project Allocation table
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCreateProjectAllocation";

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    if(!string.IsNullOrEmpty(Result))
                    {
                        string[] arrResult = Result.Split(',');
                        if (arrResult[0].ToLower() == "created")
                        {
                            ProjectAllocationID = Convert.ToInt64(arrResult[1]);
                            sqlTableName += ProjectAllocationID.ToString();
                        }
                        else
                            ErrorMessage = arrResult[0];
                    }
                }
                #endregion

                #region Create Temp Table and write input file data
                if (ProjectAllocationID > 0)
                {
                    #region Form string to create temp table
                    string strSQL = "CREATE TABLE " + sqlTableName +
                                    "(" +
                                     "[Customer Code]               VARCHAR(3)," +
                                     "[Project Code]                VARCHAR(5)," +
                                     "[Batch No.]                   VARCHAR(4)," +
                                     "[Production Emp. Code]        VARCHAR(3)," +
                                     "[QC Emp. Code]                VARCHAR(3)," +
                                     "[QA Emp. Code]                VARCHAR(3)," +
                                     "[Status]                      VARCHAR(30)," +
                                     "[Production Comments]         VARCHAR(4000)," +
                                     "[QC Comments]                 VARCHAR(4000)," +
                                     "[QA Comments]                 VARCHAR(4000)," +
                                     "[Activity]                    VARCHAR(50)";

                    #region Generate dynamic column name and data type string
                    string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                    DateTime dateValue; long lngValue; decimal decValue;
                    for (int col = 11; col <= IFws.Cells.MaxColumn; col++)
                    {
                        if (!string.IsNullOrEmpty(IFws.Cells[0, col].StringValue.Trim()))
                        {
                            ColumnName = "[" + IFws.Cells[0, col].StringValue.Trim() + "]";
                            ColumnDataType = string.Empty;
                            #region Finding out data type
                            for (int row = 1; row <= IFws.Cells.MaxRow; row++)
                            {
                                CellValue = IFws.Cells[row, col].StringValue.Trim();
                                if (!string.IsNullOrEmpty(CellValue))
                                {
                                    if (DateTime.TryParse(CellValue, out dateValue))        //if cell value is date
                                        ColumnDataType = "DATETIME";
                                    else if (Int64.TryParse(CellValue, out lngValue))         //if cell value is long
                                        ColumnDataType = "BIGINT";
                                    else if (Decimal.TryParse(CellValue, out decValue))       //if cell value is decimal
                                        ColumnDataType = "DECIMAL(18,2)";
                                    else                                                    // default string is considered
                                    {
                                        ColumnDataType = "VARCHAR(4000)";
                                        break;
                                    }
                                }
                            }
                            if (string.IsNullOrEmpty(ColumnDataType))
                                ColumnDataType = "VARCHAR(1)";
                            ColumnNameDataTypeString += "," + ColumnName + " " + ColumnDataType;
                            #endregion
                        }
                    }
                    #endregion

                    strSQL += ColumnNameDataTypeString + ")";
                    #endregion

                    using (SqlConnection conn = new SqlConnection(ConnectionString))
                    {
                        #region Creating SQL table
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = strSQL;
                        cmd.CommandType = System.Data.CommandType.Text;
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        #endregion

                        #region Writing data to SQL table
                        string excelCS = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", InputFilePath);
                        using (OleDbConnection oledbConn = new OleDbConnection(excelCS))
                        {
                            OleDbCommand oledbCommand = new OleDbCommand("SELECT * FROM [" + IFws.Name + "$]", oledbConn);
                            oledbConn.Open();
                            // Create DbDataReader to Data Worksheet  
                            DbDataReader dr = oledbCommand.ExecuteReader();
                            SqlBulkCopy bulkInsert = new SqlBulkCopy(conn);
                            bulkInsert.DestinationTableName = sqlTableName;
                            bulkInsert.WriteToServer(dr);
                        }
                        #endregion

                        var message = Request.CreateResponse(HttpStatusCode.OK);
                        conn.Close();

                        #region Move input file to Project Allocation files folder
                        DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProjectAllocationFiles/"));
                        string DestinationFileName="ProjectAllocation_" + ProjectAllocationID.ToString() + ".xlsx";
                        FileOperations.MoveFile(dirTemp, FileName, dirUploads, DestinationFileName);
                        #endregion

                        return message;
                    }
                }

                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ErrorMessage);
                #endregion
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Allocation
        [HttpGet]
        [Route("ReadProjectAllocation/{CustomerCode}/{ProjectCode}/{BatchNo}")]
        public IHttpActionResult ReadProjectAllocation(string CustomerCode, string ProjectCode, string BatchNo)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;
                ProjectAllocation projectAllocation = new ProjectAllocation();
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spViewProjectAllocation";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        projectAllocation.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectAllocation.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectAllocation.BatchNo = sqlReader["BatchNo"].ToString();
                        projectAllocation.AllocatedBy = sqlReader["AllocatedBy"].ToString();
                        projectAllocation.AllocatedOn = sqlReader["AllocatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["AllocatedOn"];
                        projectAllocation.FileName = sqlReader["FileName"].ToString();
                        projectAllocation.UniqueColumnName = sqlReader["UniqueColumnName"].ToString();
                        conn.Close();
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }

                    return Ok(projectAllocation);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Edit Allocation
        [HttpPut]
        [Route]
        public HttpResponseMessage PutProjectAllocation()           //[FromBody] ProjectAllocation projectAllocation
        {
            try
            {
                long ProjectAllocationID = 10;                                              //temporary
                string FileName = "2d65d18e-35c8-4865-98a2-026b8908c50f.xlsx";              //temporary
                string UniqueColumnName = "Material Code";                                  //temporary
                string UserID = "vic";                                                      //temporary

                //long ProjectAllocationID = projectAllocation.ProjectAllocationID;
                //string FileName = projectAllocation.FileName;      
                //string UniqueColumnName = projectAllocation.UniqueColumnName;
                //string UserID = projectAllocation.UserID;

                string sqlTableName = "ProjectAllocation_" + ProjectAllocationID.ToString();
                string ConnectionString = DBConnInfo.ConnectionString();
                string ErrorMessage = string.Empty;
                string InputFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                Workbook IFwb = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilePath);
                Worksheet IFws = IFwb.Worksheets[0];

                #region Check all values from Unique Column are unique
                int UniqueColumnNo = -1;
                for (int col = 0; col <= IFws.Cells.MaxColumn; col++)
                {
                    if (IFws.Cells[0, col].StringValue.Trim().ToLower() == UniqueColumnName.Trim().ToLower())
                    {
                        UniqueColumnNo = col;
                        break;
                    }
                }

                if (UniqueColumnNo < 0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Unique column name not found in input file");

                List<string> UniqueColumnsAllValuesList = new List<string>();
                for (int row = 1; row <= IFws.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(IFws.Cells[row, UniqueColumnNo].StringValue.Trim()))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Unique column name all values are mandatory");

                    UniqueColumnsAllValuesList.Add(IFws.Cells[row, UniqueColumnNo].StringValue.Trim());
                }

                List<string> UniqueColumnsDistinctValuesList = UniqueColumnsAllValuesList.Distinct<string>().ToList();
                if (UniqueColumnsDistinctValuesList.Count != UniqueColumnsAllValuesList.Count)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Duplicate values found in unique column");
                #endregion

                #region Form string to create temp table
                string strSQL = "CREATE TABLE " + sqlTableName +
                                "(" +
                                 "[Customer Code]               VARCHAR(3)," +
                                 "[Project Code]                VARCHAR(5)," +
                                 "[Batch No.]                   VARCHAR(4)," +
                                 "[Production Emp. Code]        VARCHAR(3)," +
                                 "[QC Emp. Code]                VARCHAR(3)," +
                                 "[QA Emp. Code]                VARCHAR(3)," +
                                 "[Status]                      VARCHAR(30)," +
                                 "[Production Comments]         VARCHAR(4000)," +
                                 "[QC Comments]                 VARCHAR(4000)," +
                                 "[QA Comments]                 VARCHAR(4000)," +
                                 "[Activity]                    VARCHAR(50)";

                #region Generate dynamic column name and data type string
                string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                DateTime dateValue; long lngValue; decimal decValue;
                for (int col = 11; col <= IFws.Cells.MaxColumn; col++)
                {
                    if (!string.IsNullOrEmpty(IFws.Cells[0, col].StringValue.Trim()))
                    {
                        ColumnName = "[" + IFws.Cells[0, col].StringValue.Trim() + "]";
                        ColumnDataType = string.Empty;
                        #region Finding out data type
                        for (int row = 1; row <= IFws.Cells.MaxRow; row++)
                        {
                            CellValue = IFws.Cells[row, col].StringValue.Trim();
                            if (!string.IsNullOrEmpty(CellValue))
                            {
                                if (DateTime.TryParse(CellValue, out dateValue))        //if cell value is date
                                    ColumnDataType = "DATETIME";
                                else if (Int64.TryParse(CellValue, out lngValue))         //if cell value is long
                                    ColumnDataType = "BIGINT";
                                else if (Decimal.TryParse(CellValue, out decValue))       //if cell value is decimal
                                    ColumnDataType = "DECIMAL(18,2)";
                                else                                                    // default string is considered
                                {
                                    ColumnDataType = "VARCHAR(4000)";
                                    break;
                                }
                            }
                        }
                        if (string.IsNullOrEmpty(ColumnDataType))
                            ColumnDataType = "VARCHAR(1)";
                        ColumnNameDataTypeString += "," + ColumnName + " " + ColumnDataType;
                        #endregion
                    }
                }
                #endregion

                strSQL += ColumnNameDataTypeString + ")";
                #endregion

                #region Update entry in Project Allocation table
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spEditProjectAllocation";

                    cmd.Parameters.AddWithValue("@ProjectAllocationID", ProjectAllocationID);
                    cmd.Parameters.AddWithValue("@UniqueColumnName", UniqueColumnName);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    if (Result.ToLower() == "updated")
                    {
                        #region Create Temp Table and write input file data

                        #region Creating SQL table
                        SqlCommand cmdCreateTable = conn.CreateCommand();
                        cmdCreateTable.CommandText = strSQL;
                        cmdCreateTable.CommandType = System.Data.CommandType.Text;
                        cmdCreateTable.ExecuteNonQuery();
                        #endregion

                        #region Writing data to SQL table
                        string excelCS = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", InputFilePath);
                        using (OleDbConnection oledbConn = new OleDbConnection(excelCS))
                        {
                            OleDbCommand oledbCommand = new OleDbCommand("SELECT * FROM [" + IFws.Name + "$]", oledbConn);
                            oledbConn.Open();
                            // Create DbDataReader to Data Worksheet  
                            DbDataReader dr = oledbCommand.ExecuteReader();
                            SqlBulkCopy bulkInsert = new SqlBulkCopy(conn);
                            bulkInsert.DestinationTableName = sqlTableName;
                            bulkInsert.WriteToServer(dr);
                        }
                        #endregion

                        #endregion

                        var message = Request.CreateResponse(HttpStatusCode.OK);
                        conn.Close();

                        #region Move input file to Project Allocation files folder
                        DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/ProjectAllocationFiles/"));
                        string DestinationFileName = "ProjectAllocation_" + ProjectAllocationID.ToString() + ".xlsx";
                        FileOperations.MoveFile(dirTemp, FileName, dirUploads, DestinationFileName);
                        #endregion

                        return message;
                    }
                    else
                    {
                        conn.Close();
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Download Project Allocation Input File
        [HttpGet]
        [Route("downloadallocationfile")]
        public HttpResponseMessage DownloadAllocationFile(string FileName)
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/Uploads/ProjectAllocationFiles/") + FileName;

                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Allocated File not found");

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
    }
}
