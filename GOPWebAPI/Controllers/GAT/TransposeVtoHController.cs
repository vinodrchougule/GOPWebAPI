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
using GOPWebAPI.Models.GAT_Models;
using Swashbuckle.Swagger;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/TransposeVtoH")]
    public class TransposeVtoHController : ApiController
    {
        #region Download Transpose Vertical To Horizontal Template File
        [HttpGet]
        [Route("DownloadTransposeVtoHFileTemplate")]
        public HttpResponseMessage DownloadTransposeVtoHFileTemplate()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileTemplatePath = HttpContext.Current.Server.MapPath("~/Templates/GAT/TransposeVtoHTemplate.xlsx");

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

        #region Validate Input File Data
        [HttpPost]
        [Route("ValidateInputFileData")]
        public HttpResponseMessage ValidateInputFileData(string FileName)
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL No' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Noun'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 3].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 4].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fifth Column with heading 'Attribute' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 5].StringValue.Trim().ToUpper() != "VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Sixth Column with heading 'Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Input file data validated Successfully.");
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

        #region Transpose the data from Vertical to Horizontal format and Download To Excel
        [HttpGet]
        [Route("TransposeDataFromVtoHAndDownloadToExcel")]
        public HttpResponseMessage TransposeDataFromVtoHAndDownloadToExcel(string FileName)
        {
            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                #region Reading data from input file
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                int maxrow = wsIF.Cells.MaxRow;
                List<TransposeVtoH> verticalItemsList = new List<TransposeVtoH>();
                for (int i = 1; i <= maxrow; i++)
                {
                    verticalItemsList.Add(new TransposeVtoH()
                    {
                        SLNO = i.ToString(),
                        Reference = wsIF.Cells[i, 1].StringValue.Trim(),
                        Noun = wsIF.Cells[i, 2].StringValue.Trim(),
                        Modifier = wsIF.Cells[i, 3].StringValue.Trim(),
                        Attribute = wsIF.Cells[i, 4].StringValue.Trim(),
                        Value = wsIF.Cells[i, 5].StringValue.Trim()
                    });
                }
                #endregion

                var ListOfVerticalItems = verticalItemsList.OrderBy(o => o.Reference).ThenBy(t => t.Noun).ThenBy(t1 => t1.Modifier);

                string OutputFileName = "Transposed Data from Vertical to Horizontal Format.xlsx";

                int RowCounter = 0;
                string pReference = "", pNoun = "", pModifier = "";
                int AttributeColNo = 4;

                #region Setting up the workbook
                AsposeHelpers asposeHelpers = new AsposeHelpers();
                Workbook wb = asposeHelpers.GetWorkbook();
                var ws = wb.Worksheets[0];
                #endregion

                #region Setting Styles
                Aspose.Cells.Style styleHeader = asposeHelpers.GetStyle(wb, 0, "header");
                Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");
                #endregion

                #region Writing column headings and setting style
                ws.Cells[0, 0].PutValue("SL No");
                ws.Cells[0, 1].PutValue("Reference");
                ws.Cells[0, 2].PutValue("Noun");
                ws.Cells[0, 3].PutValue("Modifier");
                int cntr = 1;
                for (int c = 4; c <= 203; c += 2)
                {
                    ws.Cells[0, c].PutValue("Attribute" + cntr.ToString());
                    ws.Cells[0, c + 1].PutValue("Value" + cntr.ToString());
                    cntr++;
                }

                for (int c = 0; c <= 203; c++)
                    ws.Cells[0, c].SetStyle(styleHeader);
                #endregion

                #region Writing row data
                foreach (var vi in ListOfVerticalItems)
                {
                    if (vi.Reference == pReference && vi.Noun == pNoun && vi.Modifier == pModifier)
                    {
                        AttributeColNo = AttributeColNo + 2;
                        ws.Cells[RowCounter, AttributeColNo].PutValue(vi.Attribute);
                        if (vi.Value.ToString().Length > 0)
                            ws.Cells[RowCounter, AttributeColNo + 1].PutValue(vi.Value);
                    }
                    else
                    {
                        RowCounter++;
                        AttributeColNo = 4;
                        ws.Cells[RowCounter, 0].PutValue(RowCounter);
                        ws.Cells[RowCounter, 1].PutValue(vi.Reference);
                        ws.Cells[RowCounter, 2].PutValue(vi.Noun);
                        ws.Cells[RowCounter, 3].PutValue(vi.Modifier);
                        ws.Cells[RowCounter, AttributeColNo].PutValue(vi.Attribute);
                        if (vi.Value.ToString().Length > 0)
                            ws.Cells[RowCounter, AttributeColNo + 1].PutValue(vi.Value);
                    }
                    pReference = vi.Reference;
                    pNoun = vi.Noun;
                    pModifier = vi.Modifier;
                }
                #endregion

                #region setting row data style
                for (int r = 1; r <= RowCounter; r++)
                {
                    for (int c = 0; c <= 203; c++)
                        ws.Cells[r, c].SetStyle(styleCenterAlignData);
                }
                #endregion

                //Delete input uploaded file from temp path
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                #region Saving and downloading the report
                ws.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + OutputFileName;
                wb.Save(filename);

                byte[] bytes = File.ReadAllBytes(filename);

                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                response.Content = new ByteArrayContent(bytes);

                response.Content.Headers.ContentLength = bytes.LongLength;

                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = OutputFileName;

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
