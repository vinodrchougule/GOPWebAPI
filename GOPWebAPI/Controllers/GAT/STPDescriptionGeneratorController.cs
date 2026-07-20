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
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/STPDescriptionGenerator")]
    public class STPDescriptionGeneratorController : ApiController
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

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'SL No' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 4].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Attribute' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 5].StringValue.Trim().ToUpper() != "VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

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

        #region Process And Write the output to file
        [HttpPost]
        [Route("GenerateDescriptionAndWriteToOutput")]
        public HttpResponseMessage GenerateDescriptionAndWriteToOutput(string UploadedInputFileName, string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet iws = wbIF.Worksheets[0];
                string[] SplittedAttributeNameValuePairs, SplittedAttributeNameAndValue, SplittedConsistsOfAttributeNameValuePairs;
                string GeneratedDescription = string.Empty;
                int ReferenceIDFirstRowNo = 2;
                bool IsAdditionalInformationAttributeValueBlank = false;

                iws.Cells[0, 6].PutValue("LD Output");

                for (int mRow = 1; mRow <= iws.Cells.MaxRow;)
                {
                    if (string.IsNullOrEmpty(iws.Cells[mRow, 3].StringValue.Trim()))             // if modifier is empty, no need of comma after noun
                        GeneratedDescription = iws.Cells[mRow, 2].StringValue.Trim() + "\n";
                    else
                        GeneratedDescription = iws.Cells[mRow, 2].StringValue.Trim() + "," + iws.Cells[mRow, 3].StringValue.Trim() + "\n";

                    if (!string.IsNullOrEmpty(iws.Cells[mRow, 5].StringValue.Trim()))           //first row of Reference ID
                        GeneratedDescription += iws.Cells[mRow, 4].StringValue.Trim() + ": " + iws.Cells[mRow, 5].StringValue.Trim() + "\n";

                    IsAdditionalInformationAttributeValueBlank = false;
                    for (int iRow = mRow + 1; iRow <= iws.Cells.MaxRow; iRow++)
                    {
                        if (iws.Cells[mRow, 1].StringValue.Trim().ToUpper() == iws.Cells[iRow, 1].StringValue.Trim().ToUpper())
                        {
                            if (iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "ADDITIONAL INFORMATION" &&
                                string.IsNullOrEmpty(iws.Cells[iRow, 5].StringValue.Trim()))
                                IsAdditionalInformationAttributeValueBlank = true;

                            if (!string.IsNullOrEmpty(iws.Cells[iRow, 5].StringValue.Trim()))
                            {
                                if (iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "ADDITIONAL INFORMATION" ||
                                    (iws.Cells[iRow, 4].StringValue.Trim().ToUpper().StartsWith("CONSIST OF") &&
                                     iws.Cells[iRow, 5].StringValue.Trim().Contains(';')) ||
                                    iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "REFERENCE TEXT")
                                {
                                    if (iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "ADDITIONAL INFORMATION")
                                        GeneratedDescription += "\n" + iws.Cells[iRow, 4].StringValue.Trim() + ": " + "\n";

                                    if (iws.Cells[iRow, 4].StringValue.Trim().ToUpper().StartsWith("CONSIST OF") &&
                                       iws.Cells[iRow, 5].StringValue.Trim().Contains(';'))
                                        GeneratedDescription += iws.Cells[iRow, 4].StringValue.Trim() + ": " + "\n";

                                    SplittedAttributeNameValuePairs = iws.Cells[iRow, 5].StringValue.Trim().Split(';');
                                    if (SplittedAttributeNameValuePairs.Length > 0)
                                    {
                                        if (IsAdditionalInformationAttributeValueBlank &&
                                            iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "REFERENCE TEXT") //as user requested replace REF TEXT with ADD INFO
                                            GeneratedDescription += "\n" + "ADDITIONAL INFORMATION: " + "\n";

                                        foreach (string av in SplittedAttributeNameValuePairs)
                                        {
                                            if (iws.Cells[iRow, 4].StringValue.Trim().ToUpper() == "REFERENCE TEXT")
                                                GeneratedDescription += av.Trim() + "\n";
                                            else
                                            {
                                                SplittedAttributeNameAndValue = av.Split(':');
                                                if (SplittedAttributeNameAndValue.Length == 2)
                                                {
                                                    if (!string.IsNullOrEmpty(SplittedAttributeNameAndValue[1]) &&
                                                        SplittedAttributeNameAndValue[0].Trim().ToUpper() == "CONSISTS OF")
                                                    {
                                                        GeneratedDescription += "CONSISTS OF: " + "\n";
                                                        SplittedConsistsOfAttributeNameValuePairs = SplittedAttributeNameAndValue[1].Split(';');
                                                        foreach (string coav in SplittedConsistsOfAttributeNameValuePairs)
                                                            GeneratedDescription += coav.Trim() + "\n";
                                                    }
                                                    else
                                                        GeneratedDescription += av.Trim() + "\n";
                                                }
                                                else if (SplittedAttributeNameAndValue.Length == 1)
                                                    GeneratedDescription += SplittedAttributeNameAndValue[0].Trim() + "\n";
                                            }
                                        }
                                    }
                                    else if (SplittedAttributeNameValuePairs.Length == 1)
                                        GeneratedDescription += SplittedAttributeNameValuePairs[0].Trim() + "\n";
                                }
                                else
                                    GeneratedDescription += iws.Cells[iRow, 4].StringValue.Trim() + ": " + iws.Cells[iRow, 5].StringValue.Trim() + "\n";
                            }
                            mRow++;
                        }
                        else
                        {
                            if (ReferenceIDFirstRowNo == 2)
                                iws.Cells[1, 6].PutValue(GeneratedDescription);
                            else
                                iws.Cells[ReferenceIDFirstRowNo, 6].PutValue(GeneratedDescription);
                            ReferenceIDFirstRowNo = iRow;
                            mRow = iRow;
                            break;
                        }
                    }
                    if (ReferenceIDFirstRowNo == 2)
                        iws.Cells[1, 6].PutValue(GeneratedDescription);
                    else
                        iws.Cells[ReferenceIDFirstRowNo, 6].PutValue(GeneratedDescription);
                    if (mRow == iws.Cells.MaxRow)
                        break;
                }

                #region Save the file
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

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
