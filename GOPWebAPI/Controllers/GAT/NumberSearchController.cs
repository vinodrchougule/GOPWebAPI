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
    [RoutePrefix("api/NumberSearch")]
    public class NumberSearchController : ApiController
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIFFW = wbIF.Worksheets[0];      //Input File First Worksheet

                #region Check uploaded file first worksheet has SL NO column
                if (wsIFFW.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL NO' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 1].StringValue.Trim().ToUpper() != "SEARCH")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Search' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 2].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 3].StringValue.Trim().ToUpper() != "RESULT")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Result' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validation completed successfully.");
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

        #region Search the data from input file and write the result
        [HttpPost]
        [Route("SearchDataAndWriteTheResult")]
        public HttpResponseMessage SearchDataAndWriteTheResult(string UploadedInputFileName, string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                for (int slno = 1; slno <= wsIF.Cells.MaxRow; slno++)
                    wsIF.Cells[slno, 0].PutValue(slno);
                wbIF.Save(UploadedFilepath);

                string searchText = "", descText = "", usearchText = "", resultString = "";
                int ResultColNo = 3, LastResultColNo = 3;

                for (int sRow = 1; sRow <= wsIF.Cells.MaxRow; sRow++)
                {
                    searchText = wsIF.Cells[sRow, 1].StringValue;
                    ResultColNo = 3;
                    if (!string.IsNullOrEmpty(searchText))
                    {
                        usearchText = searchText.ToUpper();
                        for (int dRow = 1; dRow <= wsIF.Cells.MaxRow; dRow++)
                        {
                            descText = wsIF.Cells[dRow, 2].StringValue;
                            if (!string.IsNullOrEmpty(descText))
                            {
                                descText = descText.ToUpper();
                                if (descText.Contains(usearchText))
                                {
                                    resultString = wsIF.Cells[dRow, ResultColNo].StringValue;
                                    while (!string.IsNullOrEmpty(resultString))
                                    {
                                        ResultColNo++;
                                        resultString = wsIF.Cells[dRow, ResultColNo].StringValue;
                                        if (ResultColNo > LastResultColNo)
                                            LastResultColNo = ResultColNo;
                                    }
                                    wsIF.Cells[dRow, ResultColNo].PutValue(searchText);
                                }
                            }
                        }
                    }
                }

                for (int c = 3; c <= LastResultColNo; c++)
                    wsIF.Cells[0, c].PutValue("Result" + (c - 3).ToString());

                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
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
    }
}
