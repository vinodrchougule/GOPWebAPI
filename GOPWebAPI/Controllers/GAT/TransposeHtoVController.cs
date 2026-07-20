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
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.DAL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using Swashbuckle.Swagger;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/TransposeHtoV")]
    public class TransposeHtoVController : ApiController
    {
        private BLLGAT _BLLGAT;

        public TransposeHtoVController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
        }

        #region Download Transpose Horizontal To Vertical Template File
        [HttpGet]
        [Route("DownloadTransposeHtoVFileTemplate")]
        public HttpResponseMessage DownloadTransposeHtoVFileTemplate()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/GAT/TransposeHtoVTemplate.xlsx");

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

        #region Validate and Write Input File Data to database
        [HttpPost]
        [Route("ValidateAndWriteInputFileDataToDatabase")]
        public HttpResponseMessage ValidateAndWriteInputFileDataToDatabase(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL NO' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'REFERENCE' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'NOUN'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 3].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'MODIFIER' not found in input file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCtr = 1; aCtr <= 200; aCtr += 2)
                {
                    hdr = wsIF.Cells[0, aCtr + 3].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 4).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' .Please download format from Template.");
                    }

                    hdr = wsIF.Cells[0, aCtr + 4].StringValue.Trim().ToUpper();
                    if (hdr != "VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 5).ToString() + "] must contain 'Value" + attNo.ToString() + "' .Please download format from Template.");
                    }
                    attNo++;
                }

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Create temp table and write uploaded file data
                string sqlTableName = _BLLGAT.CreateSQLTable("spHtoVCreateTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(FileName, sqlTableName, 0);
                #endregion

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                if (IsSucceeded)
                    return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Failed to write data to database");
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

        #region Transpose the data and Download To Excel
        [HttpGet]
        [Route("TransposeDataAndDownloadToExcel")]
        public HttpResponseMessage TransposeDataAndDownloadToExcel(string sqlTableName)
        {
            try
            {
                if (string.IsNullOrEmpty(sqlTableName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File Data not found.");

                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                DataTable dataTable = _BLLGAT.TransposeDataFromHtoV(sqlTableName);

                string FileName = "Transposed Data from Horizontal to Vertical Format.xlsx";

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
                ws.Cells[0, 2].PutValue("Noun");
                ws.Cells[0, 3].PutValue("Modifier");
                ws.Cells[0, 4].PutValue("Attribute");
                ws.Cells[0, 5].PutValue("Value");

                for (int c = 0; c <= 5; c++)
                    ws.Cells[0, c].SetStyle(styleHeader);
                #endregion

                #region Writing row data
                foreach (DataRow dr in dataTable.Rows)
                {
                    #region Writing row data
                    ws.Cells[row, 0].PutValue(dr["SLNO"].ToString().Trim());
                    ws.Cells[row, 1].PutValue(dr["Reference"].ToString().Trim());
                    ws.Cells[row, 2].PutValue(dr["Noun"] != DBNull.Value ? dr["Noun"].ToString().Trim() : "");
                    ws.Cells[row, 3].PutValue(dr["Modifier"] != DBNull.Value ? dr["Modifier"].ToString().Trim() : "");
                    ws.Cells[row, 4].PutValue(dr["Attribute"] != DBNull.Value ? dr["Attribute"].ToString().Trim() : "");
                    ws.Cells[row, 5].PutValue(dr["Value"] != DBNull.Value ? dr["Value"].ToString().Trim() : "");
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

                byte[] bytes = File.ReadAllBytes(filename);

                response.Content = new ByteArrayContent(bytes);

                response.Content.Headers.ContentLength = bytes.LongLength;

                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(filename));

                return response;
                #endregion
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