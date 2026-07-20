using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using Swashbuckle.Swagger;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.DynamicData;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/RemoveDuplicateWords")]
    public class RemoveDuplicateWordsController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet ws = wbIF.Worksheets[0];

                if (ws.Cells[0, 0].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Description' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully.");
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

        #region Remove duplicate words and Write Description to output file
        [HttpPost]
        [Route("RemoveDuplicateWordsAndWriteDescriptionToOutputFile")]
        public HttpResponseMessage RemoveDuplicateWordsAndWriteDescriptionToOutputFile(string InputFileName, string UploadedInputFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.LoadData(UploadedInputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                wsIF.Cells[0, 1].PutValue("Unique Values");
                string InputDescription = string.Empty;

                #region Removing duplicate words starts
                for (int iRow = 1; iRow <= wsIF.Cells.MaxRow; iRow++)
                {
                    InputDescription = wsIF.Cells[iRow, 0].StringValue.Trim().ToUpper();
                    string[] arrWordsFromDescription = InputDescription.Split(' ');
                    foreach (string s in arrWordsFromDescription.Distinct<string>())
                    {
                        if (s.Trim().Length > 0)
                        {
                            if (wsIF.Cells[iRow, 1].StringValue.Trim().Length > 0)
                                wsIF.Cells[iRow, 1].PutValue(wsIF.Cells[iRow, 1].StringValue.Trim() + " " + s);
                            else
                                wsIF.Cells[iRow, 1].PutValue(s);
                        }
                    }
                }
                #endregion

                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);

                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

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
