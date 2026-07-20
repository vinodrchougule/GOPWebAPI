using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Web;
using System.Web.Http;
using System.Xml.Linq;
namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/MergeMultipleWorksheets")]
    public class MergeMultipleWorksheetsController : ApiController
    {
        #region Validate the input uploaded file
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

                if (ext != ".xlsx")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");
                }

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                if (wbIF.Worksheets.Count < 2)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file with atleast 2 worksheets.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "File is valid.");
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

        #region Merge worksheets data to new worksheet and save the file
        [HttpPost]
        [Route("MergeWorksheetsDataAndSaveTheFile")]
        public HttpResponseMessage MergeWorksheetsDataAndSaveTheFile(string UploadedInputFileName, string InputFileName)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);

                wbIF.Worksheets.Add(sheetName: "MasterSheet");
                Worksheet wsMaster = wbIF.Worksheets["MasterSheet"];

                foreach(Worksheet ws in wbIF.Worksheets)
                {
                    if (ws.Name == wsMaster.Name) continue;

                    int srcLastRow = ws.Cells.MaxDataRow;
                    int srcLastCol = ws.Cells.MaxDataColumn;

                    if (srcLastCol >= 0 && srcLastRow >= 0)
                    {
                        int lastUsedRowMaster = wsMaster.Cells.MaxDataRow;
                        int pasteRow = lastUsedRowMaster + 1;

                        wsMaster.Cells.CopyRows(ws.Cells, 0, pasteRow, srcLastRow + 1);

                        for (int r = 0; r <= srcLastRow; r++)
                            wsMaster.Cells[pasteRow + r, 0].PutValue(ws.Name);
                    }
                }

                wsMaster.Move(0);
                wbIF.Worksheets.ActiveSheetIndex = 0;

                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);

                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

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
    }
}
