using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/AbbreviateFile")]
    public class AbbreviateFileController : ApiController
    {
        private BLLGAT _BLLGAT;

        public AbbreviateFileController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
        }

        #region Download Abbreviate Input File Template
        [HttpGet]
        [Route("DownloadTemplateFile")]
        public HttpResponseMessage DownloadTemplateFile(string FileName)
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/GAT/" + FileName);

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

        #region Validate the input uploaded file and Write data to server
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if(ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL No' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "INPUT DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Input Description'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 3].StringValue.Trim().ToUpper() != "ABBREVIATED DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Abbreviated Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Create temp table and write uploaded file data
                string sqlTableName = _BLLGAT.CreateSQLTable("spGATABBCreateInputFileTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(FileName, sqlTableName, 0);
                #endregion

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                if (IsSucceeded)
                    return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Failed to write input file data to database");
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate the abbreviation uploaded file and Write data to server
        [HttpPost]
        [Route("ValidateAbbreviationFile")]
        public HttpResponseMessage ValidateAbbreviationFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid abbreviation file of xlsx format only");

                Workbook wbAF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbAF.Open(UploadedFilepath);
                Worksheet wsAF = wbAF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsAF.Cells[0, 0].StringValue.Trim().ToUpper() != "EXPANDED VERSION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Expanded Version' not found in abbreviation file first worksheet. Please select a valid file.");
                }

                if (wsAF.Cells[0, 1].StringValue.Trim().ToUpper() != "ABBREVIATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Abbreviation' not found in abbreviation file first worksheet. Please select a valid file.");
                }

                if (wsAF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation File first Worksheet has no data rows.");
                }
                #endregion

                #region Create temp table and write uploaded file data
                string sqlTableName = _BLLGAT.CreateSQLTable("spGATABBCreateAbbreviationFileTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(FileName, sqlTableName, 0);
                #endregion

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                if (IsSucceeded)
                    return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Failed to write abbreviation file data to database");
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Abbreviate Input File data and write output to excel
        [HttpPost]
        [Route("AbbreviateInputFileAndWriteOutputToExcel")]
        public HttpResponseMessage AbbreviateInputFileAndWriteOutputToExcel(string UploadedInputFileName,string inputFileSqlTableName, string abbreviationFileSqlTableName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                DataTable dataTable = _BLLGAT.AbbreviateInputFile(inputFileSqlTableName, abbreviationFileSqlTableName);

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
                ws.Cells[0, 0].PutValue("SL No");
                ws.Cells[0, 1].PutValue("Reference");
                ws.Cells[0, 2].PutValue("Input Description");
                ws.Cells[0, 3].PutValue("Abbreviated Description");

                for (int c = 0; c <= 3; c++)
                    ws.Cells[0, c].SetStyle(styleHeader);
                #endregion

                #region Writing row data
                foreach (DataRow dr in dataTable.Rows)
                {
                    #region Writing row data
                    ws.Cells[row, 0].PutValue(dr["SLNO"].ToString().Trim());
                    ws.Cells[row, 1].PutValue(dr["Reference"].ToString().Trim());
                    ws.Cells[row, 2].PutValue(dr["InputDescription"] != DBNull.Value ? dr["InputDescription"].ToString().Trim() : "");
                    ws.Cells[row, 3].PutValue(dr["AbbreviatedDescription"] != DBNull.Value ? dr["AbbreviatedDescription"].ToString().Trim() : "");
                    #endregion

                    #region setting row data style
                    ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                    ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                    ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                    ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                    #endregion

                    row++;
                }
                #endregion

                ws.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wb.Save(filename);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
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

        #region Download the output file
        [HttpGet]
        [Route("DownloadOutputFile")]
        public HttpResponseMessage DownloadOutputFile(string FileFullName)
        {
            try
            {
                string FileName = Path.GetFileName(FileFullName.Replace("\\\\", "\\"));
                byte[] bytes = File.ReadAllBytes(FileFullName);

                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                response.Content = new ByteArrayContent(bytes);

                response.Content.Headers.ContentLength = bytes.LongLength;

                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileFullName));

                return response;
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
