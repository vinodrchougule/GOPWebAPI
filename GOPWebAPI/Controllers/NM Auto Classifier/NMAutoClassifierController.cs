using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;

namespace GOPWebAPI.Controllers.NM_Auto_Classifier
{
    [RoutePrefix("api/NMAutoClassifier")]
    public class NMAutoClassifierController : ApiController
    {
        private BLLNMAutoClassifier _BLLNMAutoClassifier;
        public NMAutoClassifierController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLNMAutoClassifier = new BLLNMAutoClassifier(connectionString);
        }

        #region Download File Template based on File Type
        [HttpGet]
        [Route("DownloadFileTemplate")]
        public HttpResponseMessage DownloadFileTemplate(string FileType)
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileTemplatePath = string.Empty;

                if(FileType.ToLower()=="input")
                    FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/NM Auto Classifier - Input File Template.xlsx");
                else if (FileType.ToLower() == "abbreviation")
                    FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/NM Auto Classifier - Abbreviation File Template.xlsx");
                else if (FileType.ToLower() == "noun-modifier")
                    FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/NM Auto Classifier - Noun-Modifier File Template.xlsx");
                
                //Check whether File exists.
                if (!File.Exists(FileTemplatePath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(FileTemplatePath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileTemplatePath;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileTemplatePath));

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

        #region Validate and Upload Input File Data to database
        [HttpPost]
        [Route("ValidateAndUploadInputFileData")]
        public HttpResponseMessage ValidateAndUploadInputFileData(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbIF = new Workbook();                       //Input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template 
                if(wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first worksheet has no data rows.");
                }

                if (wsIF.Cells[0, 0].StringValue.ToUpper() != "ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first worksheet must have first 'ID' column.");
                }

                if (wsIF.Cells[0, 1].StringValue.ToUpper() != "DESCRIPTION 1")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first worksheet must have second 'Description 1' column.");
                }

                if (wsIF.Cells[0, 2].StringValue.ToUpper() != "DESCRIPTION 2")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first worksheet must have third 'Description 2' column.");
                }

                if (wsIF.Cells[0, 3].StringValue.ToUpper() != "DESCRIPTION 3")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first worksheet must have fourth 'Description 3' column.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsIF.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'ID' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsIF.Cells[row, 0].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'ID' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (wsIF.Cells[row, 1].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Description 1' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }

                    if (wsIF.Cells[row, 2].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Description 2' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }

                    if (wsIF.Cells[row, 3].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Description 3' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[ID] NVARCHAR(50),";
                strSQL += "[Description1] NVARCHAR(4000),";
                strSQL += "[Description2] NVARCHAR(4000),";
                strSQL += "[Description3] NVARCHAR(4000));";
                #endregion

                #region Create temp table and write uploaded file data
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 0);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

                //Check whether File exists.
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                //return response status code
                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
            }
            catch (Exception ex)
            {
                if (sqlTableName != "tmp")
                {
                    DataFormatConverter dataFormatConverter1 = new DataFormatConverter();
                    dataFormatConverter1.DropTable(sqlTableName);
                }

                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate and Upload Abbreviation File Data to database
        [HttpPost]
        [Route("ValidateAndUploadAbbreviationFileData")]
        public HttpResponseMessage ValidateAndUploadAbbreviationFileData(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbAF = new Workbook();                       //Abbreviation file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbAF.Open(UploadedFilepath);
                Worksheet wsAF = wbAF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template 
                if (wsAF.Cells[0, 0].StringValue.ToUpper() != "SLNO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation file first worksheet must have first 'SlNo' column.");
                }

                if (wsAF.Cells[0, 1].StringValue.ToUpper() != "EXPANDED VERSION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation file first worksheet must have second 'Expanded Version' column.");
                }

                if (wsAF.Cells[0, 2].StringValue.ToUpper() != "ABBREVIATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation file first worksheet must have third 'Abbreviation' column.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsAF.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsAF.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'SlNo' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsAF.Cells[row, 0].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'SlNo' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (wsAF.Cells[row, 1].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Expanded Version' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }

                    if (wsAF.Cells[row, 2].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Abbreviation' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[SlNo] NVARCHAR(50),";
                strSQL += "[Expanded Version] NVARCHAR(4000),";
                strSQL += "[Abbreviation] NVARCHAR(4000));";
                #endregion

                #region Create temp table and write uploaded file data
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 0);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                //return response status code
                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
            }
            catch (Exception ex)
            {
                if (sqlTableName != "tmp")
                {
                    DataFormatConverter dataFormatConverter1 = new DataFormatConverter();
                    dataFormatConverter1.DropTable(sqlTableName);
                }

                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate and Upload Std. Noun-Modifier File Data to database
        [HttpPost]
        [Route("ValidateAndUploadStdNounModifierFileData")]
        public HttpResponseMessage ValidateAndUploadStdNounModifierFileData(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbNMF = new Workbook();                       //Noun-Modifier file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbNMF.Open(UploadedFilepath);
                Worksheet wsNMF = wbNMF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template 
                if (wsNMF.Cells[0, 0].StringValue.ToUpper() != "SLNO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier file first worksheet must have first 'SlNo' column.");
                }

                if (wsNMF.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier file first worksheet must have second 'Noun' column.");
                }

                if (wsNMF.Cells[0, 2].StringValue.ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier file first worksheet must have third 'Modifier' column.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNMF.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNMF.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'SlNo' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNMF.Cells[row, 0].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'SlNo' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (wsNMF.Cells[row, 1].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }

                    if (wsNMF.Cells[row, 2].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[SlNo] NVARCHAR(50),";
                strSQL += "[Noun] NVARCHAR(4000),";
                strSQL += "[Modifier] NVARCHAR(4000));";
                #endregion

                #region Create temp table and write uploaded file data
                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 0);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                //return response status code
                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
            }
            catch (Exception ex)
            {
                if (sqlTableName != "tmp")
                {
                    DataFormatConverter dataFormatConverter1 = new DataFormatConverter();
                    dataFormatConverter1.DropTable(sqlTableName);
                }

                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Process Input File and Download NM Auto Classified Report To Excel
        [HttpGet]
        [Route("ProcessAndDownloadNMAutoClassifiedReportToExcel")]
        public HttpResponseMessage ProcessAndDownloadNMAutoClassifiedReportToExcel(string InputFileTableName, string AbbreviationFileTableName, string StdNounModifierFileTableName)
        {
            try
            {
                if(string.IsNullOrEmpty(InputFileTableName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File Data not found.");

                if (string.IsNullOrEmpty(AbbreviationFileTableName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation File Data not found.");

                if (string.IsNullOrEmpty(StdNounModifierFileTableName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Std. Noun-Modifier File Data not found.");

                DataTable dataTable = new DataTable();
                dataTable = _BLLNMAutoClassifier.ProcessNMAutoClassifier(InputFileTableName, AbbreviationFileTableName, StdNounModifierFileTableName);

                if (dataTable.Rows.Count > 0)
                {
                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    string FileName = "NM Auto Classified Report.xlsx";

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
                    ws.Cells[0, 0].PutValue("ID");
                    ws.Cells[0, 1].PutValue("Description 1");
                    ws.Cells[0, 2].PutValue("Description 2");
                    ws.Cells[0, 3].PutValue("Description 3");
                    ws.Cells[0, 4].PutValue("Noun");
                    ws.Cells[0, 5].PutValue("Modifier");

                    for (int c = 0; c <= 5; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(dr["ID"].ToString().Trim());
                        ws.Cells[row, 1].PutValue(dr["Description1"].ToString().Trim());
                        ws.Cells[row, 2].PutValue(dr["Description2"] != DBNull.Value ? dr["Description2"].ToString().Trim() : "");
                        ws.Cells[row, 3].PutValue(dr["Description3"] != DBNull.Value ? dr["Description3"].ToString().Trim() : "");
                        ws.Cells[row, 4].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                        ws.Cells[row, 5].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 5].SetStyle(styleCenterAlignData);
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


    }
}
