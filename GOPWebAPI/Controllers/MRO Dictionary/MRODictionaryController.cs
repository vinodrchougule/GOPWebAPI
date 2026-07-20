using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.UI.WebControls;
using System.Xml;
using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.MRO_Dictionary;

namespace GOPWebAPI.Controllers.MRO_Dictionary
{
    [RoutePrefix("api/MRODictionary")]
    public class MRODictionaryController : ApiController
    {
        #region Upload MRO Dictionary
        #region Download MRO Dictionary Template
        [HttpGet]
        [Route("DownloadMRODictionaryTemplate")]
        public HttpResponseMessage DownloadMRODictionaryTemplate()
        {
            try
            {
                string FileName = "MRO Dictionary Upload Template.xlsx";

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/Templates/" + FileName);

                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");

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

        #region Validate the correct file upload
        [HttpGet]
        [Route("ValidateCorrectFileUpload")]
        public HttpResponseMessage ValidateCorrectFileUpload(string FileName)
        {
            try
            {
                #region Check uploaded file exists in temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (!File.Exists(dirTemp + FileName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found. Please upload valid MRO Dictionary file");
                #endregion

                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
                Workbook wbIF = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);

                if (wbIF.Worksheets.Count != 6)
                {
                    File.Delete(InputFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file should have 6 worksheets as per template.");
                }

                if (wbIF.Worksheets[0].Cells.MaxRow == 0)
                {
                    File.Delete(InputFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Uploaded file has no data in Noun worksheet");
                }

                //return response status code
                return Request.CreateResponse(HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read the count of all data from file
        [HttpGet]
        [Route("ReadCountOfAllDataFromFile")]
        public IHttpActionResult ReadCountOfAllDataFromFile(string FileName)
        {
            try
            {
                MRODictionaryFileDataCounts mroDictionaryFileDataCounts= new MRODictionaryFileDataCounts();

                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                Workbook wbIF = new Workbook();                       //input file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);

                mroDictionaryFileDataCounts.NoOfNouns = wbIF.Worksheets[0].Cells.MaxRow;
                mroDictionaryFileDataCounts.NoOfNounModifiers = wbIF.Worksheets[1].Cells.MaxRow;
                mroDictionaryFileDataCounts.NoOfNounSynonyms = wbIF.Worksheets[2].Cells.MaxRow;
                mroDictionaryFileDataCounts.NoOfNounModifierAttributes = wbIF.Worksheets[3].Cells.MaxRow;
                mroDictionaryFileDataCounts.NoOfNounModifierAttributeEVVs = wbIF.Worksheets[4].Cells.MaxRow;
                mroDictionaryFileDataCounts.NoOfNounModifiersMappedToUNSPSC = wbIF.Worksheets[5].Cells.MaxRow;

                //return list to the request
                return Ok(mroDictionaryFileDataCounts);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate and Upload Noun Definitions to database
        [HttpPost]
        [Route("ValidateAndUploadNounDefinitions")]
        public HttpResponseMessage ValidateAndUploadNounDefinitions(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNoun = wbUF.Worksheets[0];

                #region Check uploaded file Noun worksheet has columns as per template 
                if (wsNoun.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNoun.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNoun.Cells[0, 2].StringValue.ToUpper() != "NOUN DEFINITION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun worksheet must have third 'Noun Definition' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNoun.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNoun.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNoun.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNoun.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNoun.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (wsNoun.Cells[row, 2].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun Definition' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Noun Definition] VARCHAR(4000));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath,0);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Upload Noun-Modifier Definitions/Guidelines to database
        [HttpPost]
        [Route("ValidateAndUploadNounModifierDefinitions")]
        public HttpResponseMessage ValidateAndUploadNounModifierDefinitions(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNoun = wbUF.Worksheets[0];
                Worksheet wsNounModifier = wbUF.Worksheets[1];

                #region Check uploaded file Noun worksheet has columns as per template 
                if (wsNounModifier.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNounModifier.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNounModifier.Cells[0, 2].StringValue.ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier worksheet must have third 'Modifier' column. Please select valid file.");
                }

                if (wsNounModifier.Cells[0, 3].StringValue.ToUpper() != "NOUN - MODIFIER DEFINITION / GUIDELINES")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Modifier worksheet must have fourth 'Noun - Modifier Definition / Guidelines' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNounModifier.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNounModifier.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounModifier.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifier.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounModifier.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifier.Cells[row, 2].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounModifier.Cells[row, 2].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot exceed 100 characters. Row No.:" + row.ToString());
                    }

                    if (wsNounModifier.Cells[row, 3].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun - Modifier Definition / Guidelines' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Modifier] VARCHAR(100),";
                strSQL += "[Noun - Modifier Definition / Guidelines] VARCHAR(4000));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 1);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Upload Noun-Synonyms to database
        [HttpPost]
        [Route("ValidateAndUploadNounSynonyms")]
        public HttpResponseMessage ValidateAndUploadNounSynonyms(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNoun = wbUF.Worksheets[0];
                Worksheet wsNounSynonym = wbUF.Worksheets[2];

                #region Check uploaded file Noun Synonym worksheet has columns as per template 
                if (wsNounSynonym.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Synonym worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNounSynonym.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Synonym worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNounSynonym.Cells[0, 2].StringValue.ToUpper() != "SYNONYM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Synonym worksheet must have third 'Synonym' column. Please select valid file.");
                }

                if (wsNounSynonym.Cells[0, 3].StringValue.ToUpper() != "SYNONYM DEFINITION / GUIDELINES")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Noun-Synonym worksheet must have fourth 'Synonym Definition / Guidelines' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNounSynonym.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNounSynonym.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounSynonym.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounSynonym.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounSynonym.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounSynonym.Cells[row, 2].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Synonym' column value cannot be empty/blank. Row No.:" + row.ToString());
                    }

                    if (wsNounSynonym.Cells[row, 2].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Synonym' column value cannot exceed 100 characters. Row No.:" + row.ToString());
                    }

                    if (wsNounSynonym.Cells[row, 3].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Synonym Definition / Guidelines' column value cannot exceed 4000 characters. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Synonym] VARCHAR(100),";
                strSQL += "[Synonym Definition / Guidelines] VARCHAR(4000));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 2);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Upload Noun-Modifier Attributes to database
        [HttpPost]
        [Route("ValidateAndUploadNounModifierAttributes")]
        public HttpResponseMessage ValidateAndUploadNounModifierAttributes(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNounModifier = wbUF.Worksheets[1];
                Worksheet wsNounModifierAttribute = wbUF.Worksheets[3];

                #region Check uploaded file Noun-Modifier Attribute worksheet has columns as per template 
                if (wsNounModifierAttribute.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 2].StringValue.ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have third 'Modifier' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 3].StringValue.ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have fourth 'Attribute' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 4].StringValue.ToUpper() != "PRIORITY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have fifth 'Priority' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 5].StringValue.ToUpper() != "MANDATORY / OPTIONAL")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have sixth 'Mandatory / Optional' column. Please select valid file.");
                }

                if (wsNounModifierAttribute.Cells[0, 6].StringValue.ToUpper() != "ATTRIBUTE GUIDELINES")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute' worksheet must have seventh 'Attribute Guidelines' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNounModifierAttribute.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 2].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 2].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot exceed 100 characters in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 3].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Attribute' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 3].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Attribute' column value cannot exceed 100 characters in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 4].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Priority' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 4].StringValue.Length > 1)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Priority' column value cannot exceed 1 character in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttribute.Cells[row, 5].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mandatory / Optional' column value cannot be empty/blank in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 5].StringValue.Length > 1)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mandatory / Optional' column value cannot exceed 1 character in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttribute.Cells[row, 6].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Attribute Guidelines' column value cannot exceed 4000 characters in Noun-Modifier Attribute worksheet. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Modifier] VARCHAR(100),";
                strSQL += "[Attribute] VARCHAR(100),";
                strSQL += "[Priority] VARCHAR(1),";
                strSQL += "[Mandatory / Optional] VARCHAR(1),";
                strSQL += "[Attribute Guidelines] VARCHAR(4000));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 3);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Upload Noun-Modifier Attribute EVVs to database
        [HttpPost]
        [Route("ValidateAndUploadNounModifierAttributeEVVs")]
        public HttpResponseMessage ValidateAndUploadNounModifierAttributeEVVs(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNounModifierAttribute = wbUF.Worksheets[3];
                Worksheet wsNounModifierAttributeEVVs = wbUF.Worksheets[4];

                #region Check uploaded file Noun-Modifier Attribute EVVs worksheet has columns as per template 
                if (wsNounModifierAttributeEVVs.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNounModifierAttributeEVVs.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNounModifierAttributeEVVs.Cells[0, 2].StringValue.ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have third 'Modifier' column. Please select valid file.");
                }

                if (wsNounModifierAttributeEVVs.Cells[0, 3].StringValue.ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have fourth 'Attribute' column. Please select valid file.");
                }

                if (wsNounModifierAttributeEVVs.Cells[0, 4].StringValue.ToUpper() != "ENUMERATED VALID VALUE (EVV)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have fifth 'Enumerated Valid Value (EVV)' column. Please select valid file.");
                }

                if (wsNounModifierAttributeEVVs.Cells[0, 5].StringValue.ToUpper() != "PRIORITY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun-Modifier Attribute EVVs' worksheet must have sixth 'Priority' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNounModifierAttributeEVVs.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 2].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 2].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot exceed 100 characters in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 3].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Attribute' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 3].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Attribute' column value cannot exceed 100 characters in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 4].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Enumerated Valid Value (EVV)' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 4].StringValue.Length > 4000)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Enumerated Valid Value (EVV)' column value cannot exceed 4000 characters in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierAttributeEVVs.Cells[row, 5].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Priority' column value cannot be empty/blank in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierAttributeEVVs.Cells[row, 5].StringValue.Length > 1)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Priority' column value cannot exceed 1 character in Noun-Modifier Attribute EVVs worksheet. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Modifier] VARCHAR(100),";
                strSQL += "[Attribute] VARCHAR(100),";
                strSQL += "[Enumerated Valid Value (EVV)] VARCHAR(4000),";
                strSQL += "[Priority] VARCHAR(1));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 4);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Upload Noun-Modifier Mapped UNSPSCs  to database
        [HttpPost]
        [Route("ValidateAndUploadNounModifierMappedUNSPSCs")]
        public HttpResponseMessage ValidateAndUploadNounModifierMappedUNSPSCs(string FileName)
        {
            string sqlTableName = "tmp";

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
                Workbook wbUF = new Workbook();                       //Uploaded file work book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUF.Open(UploadedFilepath);
                Worksheet wsNounModifier = wbUF.Worksheets[1];
                Worksheet wsNounModifierUNSPSC = wbUF.Worksheets[5];

                #region Check uploaded file Noun-Modifier Mapped UNSPSC worksheet has columns as per template
                if (wsNounModifierUNSPSC.Cells[0, 0].StringValue.ToUpper() != "SL. NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have first 'Sl. No.' column. Please select valid file.");
                }

                if (wsNounModifierUNSPSC.Cells[0, 1].StringValue.ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have second 'Noun' column. Please select valid file.");
                }

                if (wsNounModifierUNSPSC.Cells[0, 2].StringValue.ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have third 'Modifier' column. Please select valid file.");
                }

                if (wsNounModifierUNSPSC.Cells[0, 3].StringValue.ToUpper() != "UNSPSC VERSION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have fourth 'UNSPSC Version' column. Please select valid file.");
                }

                if (wsNounModifierUNSPSC.Cells[0, 4].StringValue.ToUpper() != "UNSPSC CODE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have fifth 'UNSPSC Code' column. Please select valid file.");
                }

                if (wsNounModifierUNSPSC.Cells[0, 5].StringValue.ToUpper() != "UNSPSC CATEGORY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Mapped UNSPSC' worksheet must have sixth 'UNSPSC Category' column. Please select valid file.");
                }
                #endregion

                #region Check column values length exceeds max length or value is empty/blank
                for (int row = 1; row <= wsNounModifierUNSPSC.Cells.MaxRow; row++)
                {
                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 0].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 0].StringValue.Length > 10)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Sl. No.' column value cannot exceed 10 characters in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 1].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 1].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column value cannot exceed 50 characters in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 2].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 2].StringValue.Length > 100)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column value cannot exceed 100 characters in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 3].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Version' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 3].StringValue.Length > 50)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Version' column value cannot exceed 50 characters in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 4].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Code' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 4].StringValue.Length != 8)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Code' column value must be 8 digits code in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (string.IsNullOrEmpty(wsNounModifierUNSPSC.Cells[row, 5].StringValue))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Category' column value cannot be empty/blank in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }

                    if (wsNounModifierUNSPSC.Cells[row, 5].StringValue.Length > 255)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Category' column value cannot exceed 255 characters in Noun-Modifier Mapped UNSPSC worksheet. Row No.:" + row.ToString());
                    }
                }
                #endregion

                #region Form string to create temp table
                sqlTableName = "tbl" + Guid.NewGuid().ToString().Replace("-", "");    //SQL table name must start with alphabet and should not contain hyphen
                string strSQL = "CREATE TABLE " + sqlTableName + "(";

                strSQL += "[Sl.No.] VARCHAR(10),";
                strSQL += "[Noun] VARCHAR(50),";
                strSQL += "[Modifier] VARCHAR(100),";
                strSQL += "[UNSPSC Version] VARCHAR(50),";
                strSQL += "[UNSPSC Code] NVARCHAR(8),";
                strSQL += "[UNSPSC Category] NVARCHAR(255));";
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
                DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, 5);
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                bulkcopy.DestinationTableName = sqlTableName;
                bulkcopy.WriteToServer(excelData);
                #endregion
                #endregion

                conn.Close();

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

        #region Validate and Update MRO Dictionary
        [HttpPost]
        [Route("ValidateAndUpdateMRODictionary")]
        public HttpResponseMessage ValidateAndUpdateMRODictionary([FromBody] MRODictionaryUploadModel model)
        {
            try
            {
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());

                #region Update MRO Dictionary
                //Initialize command object
                SqlCommand cmdMD = new SqlCommand();
                cmdMD.Connection = conn;
                cmdMD.CommandType = CommandType.StoredProcedure;
                cmdMD.CommandText = "spMRODictionaryCreate";

                #region Adding Stored Procedure Parameters
                cmdMD.Parameters.AddWithValue("@UploadedFileName", model.UploadedFileName);
                cmdMD.Parameters.AddWithValue("@VersionNameOrNo", model.VersionNameOrNo);
                cmdMD.Parameters.AddWithValue("@NounTableName", model.NounTableName);
                cmdMD.Parameters.AddWithValue("@NounModifierTableName", model.NounModifierTableName);
                cmdMD.Parameters.AddWithValue("@NounSynonymTableName", model.NounSynonymTableName);
                cmdMD.Parameters.AddWithValue("@NounModifierAttributeTableName", model.NounModifierAttributeTableName);
                cmdMD.Parameters.AddWithValue("@NounModifierAttributeValuesTableName", model.NounModifierAttributeValuesTableName);
                cmdMD.Parameters.AddWithValue("@NounModifierMappedUNSPSCs", model.NounModifierMappedUNSPSCs);
                cmdMD.Parameters.AddWithValue("@UserID", model.UserID);
                #endregion

                //Calling sp to update MRO Dictionary
                cmdMD.CommandTimeout = 0;
                conn.Open();
                string Result = cmdMD.ExecuteScalar().ToString();
                conn.Close();

                string[] arrResult = Result.Split('_');
                if (arrResult[0].Trim().ToLower() == "created")
                {
                    int MRODictionaryID = Convert.ToInt32(arrResult[1]);
                    string NewMRODictionaryFileName = MRODictionaryID.ToString() + '-' + model.UploadedFileName;

                    #region Input File move Starts
                    if (File.Exists(dirTemp + model.UploadedTempFileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/"));
                        FileOperations.MoveFile(dirTemp, model.UploadedTempFileName, dirUploads, NewMRODictionaryFileName);
                    }
                    #endregion

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, MRODictionaryID);
                }
                else
                {
                    if (File.Exists(dirTemp + model.UploadedTempFileName))
                        File.Delete(dirTemp + model.UploadedTempFileName);

                    //return error response status code
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, arrResult[0]);
                }
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

        #region Read MRO Dictionaries List
        [HttpGet]
        [Route("ReadMRODictionariesList")]
        public IHttpActionResult ReadMRODictionariesList()
        {
            try
            {
                //Create a list to hold the MRO Dictionary File Data Counts
                List<MRODictionaryFileDataCounts> mroDictionaryFileDataCountsList = new List<MRODictionaryFileDataCounts>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionariesList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of MRO Dictionary File Data Counts
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        MRODictionaryFileDataCounts mroDictionaryFileDataCounts= new MRODictionaryFileDataCounts();
                        mroDictionaryFileDataCounts.MRODictionaryID = Convert.ToInt32(sqlReader["MRODictionaryID"]);
                        mroDictionaryFileDataCounts.UpdatedOn = Convert.ToDateTime(sqlReader["UpdatedOn"]);
                        mroDictionaryFileDataCounts.UpdatedBy = sqlReader["UpdatedBy"].ToString();
                        mroDictionaryFileDataCounts.VersionNameOrNo = sqlReader["VersionNameOrNo"].ToString();
                        mroDictionaryFileDataCounts.UploadedFileName = sqlReader["UploadedFileName"].ToString();
                        mroDictionaryFileDataCounts.NoOfNouns = Convert.ToInt32(sqlReader["NoOfNouns"]);
                        mroDictionaryFileDataCounts.NoOfNounModifiers = Convert.ToInt32(sqlReader["NoOfNounModifiers"]);
                        mroDictionaryFileDataCounts.NoOfNounSynonyms = Convert.ToInt32(sqlReader["NoOfNounSynonyms"]);
                        mroDictionaryFileDataCounts.NoOfNounModifierAttributes = Convert.ToInt32(sqlReader["NoOfNounModifierAttributes"]);
                        mroDictionaryFileDataCounts.NoOfNounModifierAttributeEVVs = Convert.ToInt32(sqlReader["NoOfNounModifierAttributeEVVs"]);
                        mroDictionaryFileDataCounts.NoOfNounModifiersMappedToUNSPSC = Convert.ToInt32(sqlReader["NoOfNounModifiersMappedToUNSPSC"]);
                        mroDictionaryFileDataCountsList.Add(mroDictionaryFileDataCounts);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(mroDictionaryFileDataCountsList);
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

        #region Download Selected MRO Dictionary Version Data file
        [HttpGet]
        [Route("DownloadSelectedMRODictionaryVersionDataFile")]
        public HttpResponseMessage DownloadSelectedMRODictionaryVersionDataFile(int MRODictionaryID, string FileName, string UserID)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;
                int LastRowNo = 0;

                if (!AccessControl.CanUserAccessPage(UserID, "Download MRO Dictionary"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/") + MRODictionaryID.ToString() + '-' + FileName;

                if (!File.Exists(filePath))
                {
                    //Check whether File exists.
                    if (!File.Exists(filePath))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");
                }

                #region Open the workbook to write the data
                Workbook wb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wb.Open(filePath);
                var wsNoun = wb.Worksheets[0];
                var wsNounModifier = wb.Worksheets[1];
                var wsNounSynonym = wb.Worksheets[2];
                var wsNounModifierAttribute = wb.Worksheets[3];
                var wsNounModifierAttributeValues = wb.Worksheets[4];
                var wsNounModifierUNSPSC = wb.Worksheets[5];

                #region Setting Styles
                Aspose.Cells.Style styleCenterAlignData = wsNoun.Cells[0, 3].GetStyle();
                Aspose.Cells.Style styleLeftAlignData = wsNoun.Cells[0, 3].GetStyle();

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
                #endregion

                #region Write the Not Imported Noun Details to file
                LastRowNo = wsNoun.Cells.MaxRow;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNoun.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNoun.Cells[LastRowNo,0].StringValue.Trim()) + 1);
                        wsNoun.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNoun.Cells[LastRowNo + 1, 2].PutValue(sqlReader["NounDefinition"].ToString());
                        #endregion

                        #region setting row data style
                        wsNoun.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNoun.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNoun.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Write the Not Imported Noun-Modifier Details to file
                LastRowNo = wsNounModifier.Cells.MaxRow;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounModifierDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun-Modifier Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNounModifier.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNounModifier.Cells[LastRowNo, 0].StringValue.Trim()) + 1);
                        wsNounModifier.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNounModifier.Cells[LastRowNo + 1, 2].PutValue(sqlReader["Modifier"].ToString());
                        wsNounModifier.Cells[LastRowNo + 1, 3].PutValue(sqlReader["NounModifierDefinitionOrGuidelines"].ToString());
                        #endregion

                        #region setting row data style
                        wsNounModifier.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNounModifier.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNounModifier.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        wsNounModifier.Cells[LastRowNo + 1, 3].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Write the Not Imported Noun Synonym Details to file
                LastRowNo = wsNounSynonym.Cells.MaxRow;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounSynonymDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun-Modifier Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNounSynonym.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNounSynonym.Cells[LastRowNo, 0].StringValue.Trim()) + 1);
                        wsNounSynonym.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNounSynonym.Cells[LastRowNo + 1, 2].PutValue(sqlReader["Synonym"].ToString());
                        wsNounSynonym.Cells[LastRowNo + 1, 3].PutValue(sqlReader["SynonymDefinitionOrGuidelines"].ToString());
                        #endregion

                        #region setting row data style
                        wsNounSynonym.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNounSynonym.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNounSynonym.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        wsNounSynonym.Cells[LastRowNo + 1, 3].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Write the Not Imported Noun-Modifier Attribute Details to file
                LastRowNo = wsNounModifierAttribute.Cells.MaxRow;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounModifierAttributeDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun-Modifier Attribute Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNounModifierAttribute.Cells[LastRowNo, 0].StringValue.Trim()) + 1);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 2].PutValue(sqlReader["Modifier"].ToString());
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 3].PutValue(sqlReader["Attribute"].ToString());
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 4].PutValue(sqlReader["Priority"].ToString());
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 5].PutValue(sqlReader["MandatoryOrOptional"].ToString());
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 6].PutValue(sqlReader["AttributeGuidelines"].ToString());
                        #endregion

                        #region setting row data style
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 3].SetStyle(styleLeftAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 4].SetStyle(styleLeftAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 5].SetStyle(styleLeftAlignData);
                        wsNounModifierAttribute.Cells[LastRowNo + 1, 6].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Write the Not Imported Noun-Modifier Attribute Values Details to file
                LastRowNo = wsNounModifierAttributeValues.Cells.MaxRow;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounModifierAttributeValuesDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun-Modifier Attribute Values Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNounModifierAttributeValues.Cells[LastRowNo, 0].StringValue.Trim()) + 1);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 2].PutValue(sqlReader["Modifier"].ToString());
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 3].PutValue(sqlReader["Attribute"].ToString());
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 4].PutValue(sqlReader["EnumeratedValidValue"].ToString());
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 5].PutValue(sqlReader["Priority"].ToString());
                        #endregion

                        #region setting row data style
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 3].SetStyle(styleLeftAlignData);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 4].SetStyle(styleLeftAlignData);
                        wsNounModifierAttributeValues.Cells[LastRowNo + 1, 5].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Write the Not Imported Noun-Modifier UNSPSC Details to file
                LastRowNo = wsNounModifierUNSPSC.Cells.MaxRow;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryGetNotImportedNounModifierUNSPSCDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", MRODictionaryID);

                    //Call sp to get all Not Imported Noun-Modifier UNSPSC Details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        #region Writing row data
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 0].PutValue(Convert.ToInt32(wsNounModifierUNSPSC.Cells[LastRowNo, 0].StringValue.Trim()) + 1);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 1].PutValue(sqlReader["Noun"].ToString());
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 2].PutValue(sqlReader["Modifier"].ToString());
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 3].PutValue(sqlReader["Scheme_Name"].ToString());
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 4].PutValue(sqlReader["Code"].ToString());
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 5].PutValue(sqlReader["Category"].ToString());
                        #endregion

                        #region setting row data style
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 0].SetStyle(styleCenterAlignData);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 1].SetStyle(styleLeftAlignData);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 2].SetStyle(styleLeftAlignData);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 3].SetStyle(styleCenterAlignData);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 4].SetStyle(styleLeftAlignData);
                        wsNounModifierUNSPSC.Cells[LastRowNo + 1, 5].SetStyle(styleLeftAlignData);
                        #endregion

                        LastRowNo++;
                    }
                    conn.Close();
                }
                #endregion

                #region Saving and downloading the file
                wsNoun.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + FileName;
                wb.Save(filename);

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(filename);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = filename;

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

        #region Delete the Selected MRO Dictionary Version Data
        [HttpPatch]
        [Route("DeleteSelectedMRODictionaryVersion")]
        public HttpResponseMessage DeleteSelectedMRODictionaryVersion(int id, string FileName, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete MRO Dictionary"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spMRODictionaryDelete";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@MRODictionaryID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete MRO Dictionary Version
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "deleted")
                    {
                        #region Delete MRO Dictionary Version Data File
                        string UploadedFileName = id.ToString() + '-' + FileName.Trim().ToUpper();
                        DirectoryInfo dirMRODictionaryUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/"));
                        foreach (FileInfo file in dirMRODictionaryUploads.GetFiles())
                        {
                            if (file.Name.ToUpper() == UploadedFileName)
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

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
        #endregion
    }
}
