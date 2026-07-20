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
    [RoutePrefix("api/OEMDescriptionGenerator")]
    public class OEMDescriptionGeneratorController : ApiController
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
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Reference ID' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "NM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'NM' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "CLASS_NAME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Class_Name' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "VENDOR/VENDOR CODE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Vendor/Vendor Code' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 4].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 5].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 6].StringValue.Trim().ToUpper() != "SEQ")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [G1] with Value 'Seq' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 7].StringValue.Trim().ToUpper() != "ATTRIBUTE DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [H1] with Value 'Attribute Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 8].StringValue.Trim().ToUpper() != "FLAG")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [I1] with Value 'Flag' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 9].StringValue.Trim().ToUpper() != "ITEM VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [J1] with Value 'Item Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                for(int c = 0; c <= 9; c++)
                {
                    if (ws2.Cells[0, c].StringValue.Trim().ToUpper() != ws1.Cells[0, c].StringValue.Trim().ToUpper())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [1," + (c+1).ToString() + "] with Value '" + ws1.Cells[0, c].StringValue.Trim() + "' not found in input file second worksheet. Please select a valid file.");
                    }
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
        [Route("GenerateOEMDescriptionAndWriteToOutput")]
        public HttpResponseMessage GenerateOEMDescriptionAndWriteToOutput(string UploadedInputFileName, string InputFileName, int MaxChars)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet iws = wbIF.Worksheets[0];
                Worksheet ows = wbIF.Worksheets[1];

                #region Create a list of Additional Features EN attribute distinct rows
                List<AdditionalFeatures> AdditionalFeaturesList = new List<AdditionalFeatures>();
                int inputMaxRows = iws.Cells.MaxRow;
                for (int ir = 1; ir <= inputMaxRows; ir++)
                {
                    if ((ir == inputMaxRows && iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en") ||
                        (iws.Cells[ir, 0].StringValue.Trim().ToLower() != iws.Cells[ir + 1, 0].StringValue.Trim().ToLower() &&
                         iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en") ||
                        (iws.Cells[ir, 0].StringValue.Trim().ToLower() == iws.Cells[ir + 1, 0].StringValue.Trim().ToLower() &&
                         iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en" &&
                         iws.Cells[ir + 1, 7].StringValue.Trim().ToLower() != "additional features en"))
                    {
                        AdditionalFeatures objAdditionalFeatures = new AdditionalFeatures();
                        objAdditionalFeatures.ReferenceID = iws.Cells[ir, 0].StringValue.Trim();
                        objAdditionalFeatures.NM = iws.Cells[ir, 1].StringValue.Trim();
                        objAdditionalFeatures.ClassName = iws.Cells[ir, 2].StringValue.Trim();
                        objAdditionalFeatures.VendorCode = iws.Cells[ir, 3].StringValue.Trim();
                        objAdditionalFeatures.Noun = iws.Cells[ir, 4].StringValue.Trim();
                        objAdditionalFeatures.Modifier = iws.Cells[ir, 5].StringValue.Trim();
                        objAdditionalFeatures.Seq = Convert.ToInt32(iws.Cells[ir, 6].StringValue.Trim());
                        objAdditionalFeatures.Flag = iws.Cells[ir, 8].StringValue.Trim();
                        AdditionalFeaturesList.Add(objAdditionalFeatures);
                    }
                }
                #endregion

                string AttributeNameValueString = string.Empty;
                int LastRowFromOutputWroksheet = 1;
                bool IsItLastAdditionalFeaturesAttribute = false;
                bool IsAdditionalFeaturesENAttributeProcessed = false;
                List<string> MeaningfulSplittedString = new List<string>();

                foreach (AdditionalFeatures af in AdditionalFeaturesList)
                {
                    IsItLastAdditionalFeaturesAttribute = false;
                    AttributeNameValueString = string.Empty;
                    IsAdditionalFeaturesENAttributeProcessed = false;
                    for (int ir = 1; ir <= inputMaxRows; ir++)     //Processing from 1st Row to Last Row of Input Worksheet
                    {
                        if (iws.Cells[ir, 0].StringValue.Trim().ToLower() == af.ReferenceID.Trim().ToLower())       //if Ref ID from worksheet matches to Ref ID from List
                        {
                            if (IsAdditionalFeaturesENAttributeProcessed)
                            {
                                if ((iws.Cells[ir, 7].StringValue.Trim().ToLower() != "equipment name" && iws.Cells[ir, 7].StringValue.Trim().ToLower() != "equipment type") ||
                                    ((iws.Cells[ir, 7].StringValue.Trim().ToLower() == "equipment name") && iws.Cells[ir, 9].StringValue.Trim().Length <= MaxChars) ||
                                    ((iws.Cells[ir, 7].StringValue.Trim().ToLower() == "equipment type") && iws.Cells[ir, 9].StringValue.Trim().Length <= MaxChars))
                                {
                                    #region write static rows
                                    LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(iws.Cells[ir, 0].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(iws.Cells[ir, 1].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(iws.Cells[ir, 2].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(iws.Cells[ir, 6].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 7].PutValue(iws.Cells[ir, 7].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 8].PutValue(iws.Cells[ir, 8].StringValue.Trim());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 9].PutValue(iws.Cells[ir, 9].StringValue.Trim());
                                    #endregion
                                }
                                else
                                {
                                    if (iws.Cells[ir, 9].StringValue.Trim().StartsWith("01)"))
                                        MeaningfulSplittedString = SplitInputString(iws.Cells[ir, 9].StringValue.Trim().Substring(3), MaxChars);
                                    else
                                        MeaningfulSplittedString = SplitInputString(iws.Cells[ir, 9].StringValue.Trim(), MaxChars);
                                    foreach (string ms in MeaningfulSplittedString)
                                    {
                                        #region Write splitted meaningful strings to the output worksheet
                                        LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(iws.Cells[ir, 0].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(iws.Cells[ir, 1].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(iws.Cells[ir, 2].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(iws.Cells[ir, 6].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 7].PutValue(iws.Cells[ir, 7].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 8].PutValue(iws.Cells[ir, 8].StringValue.Trim());
                                        ows.Cells[LastRowFromOutputWroksheet + 1, 9].PutValue(ms);
                                        #endregion
                                    }
                                }
                            }

                            #region Form Attribute Name Value string
                            if (!IsItLastAdditionalFeaturesAttribute)
                            {
                                if (string.IsNullOrEmpty(AttributeNameValueString))
                                {
                                    if (iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en")
                                        AttributeNameValueString = iws.Cells[ir, 9].StringValue.Trim();
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(iws.Cells[ir, 9].StringValue.Trim()))
                                            AttributeNameValueString = iws.Cells[ir, 7].StringValue.Trim() + ":" + iws.Cells[ir, 9].StringValue.Trim();
                                    }
                                }
                                else
                                {
                                    if (iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en")
                                    {
                                        if (!string.IsNullOrEmpty(iws.Cells[ir, 9].StringValue.Trim()))
                                            AttributeNameValueString += "," + iws.Cells[ir, 9].StringValue.Trim();
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(iws.Cells[ir, 9].StringValue.Trim()))
                                            AttributeNameValueString += "," + iws.Cells[ir, 7].StringValue.Trim() + ":" + iws.Cells[ir, 9].StringValue.Trim();
                                    }
                                }
                            }
                            #endregion

                            if ((ir == inputMaxRows && iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en") ||
                                (iws.Cells[ir, 0].StringValue.Trim().ToLower() != iws.Cells[ir + 1, 0].StringValue.Trim().ToLower() &&
                                 iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en") ||
                                (iws.Cells[ir, 0].StringValue.Trim().ToLower() == iws.Cells[ir + 1, 0].StringValue.Trim().ToLower() &&
                                 iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en" &&
                                 iws.Cells[ir + 1, 7].StringValue.Trim().ToLower() != "additional features en"))
                            {
                                IsItLastAdditionalFeaturesAttribute = true;
                            }

                            if (IsItLastAdditionalFeaturesAttribute && !IsAdditionalFeaturesENAttributeProcessed)
                            {
                                //split AttributeNameValueString and write it to output worksheet
                                MeaningfulSplittedString = SplitInputString(AttributeNameValueString, MaxChars);
                                foreach (string ms in MeaningfulSplittedString)
                                {
                                    #region Write splitted meaningful strings to the output worksheet
                                    LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(af.ReferenceID);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(af.NM);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(af.ClassName);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(af.VendorCode);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(af.Noun);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(af.Modifier);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(af.Seq.ToString());
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 7].PutValue("ADDITIONAL FEATURES EN");
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 8].PutValue(af.Flag);
                                    ows.Cells[LastRowFromOutputWroksheet + 1, 9].PutValue(ms);
                                    #endregion
                                }
                                IsAdditionalFeaturesENAttributeProcessed = true;
                            }
                        }
                    }
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

        private List<string> SplitInputString(string inputString, int MaxCharacters = 70)
        {
            List<string> SplittedStringList = new List<string>();
            string MaxCharactersMeaningfulString = string.Empty, AttributeValueSlNoString = "01)";
            string InputStringToSplit = inputString.Trim(), SpaceSplittedString = string.Empty, SpaceCombinedString = string.Empty;
            int AttributeValueSlNo = 1;

            if (string.IsNullOrEmpty(InputStringToSplit) || InputStringToSplit.IndexOf(',') < 0)
            {
                if (string.IsNullOrEmpty(InputStringToSplit))
                {
                    SplittedStringList.Add(InputStringToSplit);
                    return SplittedStringList;
                }
                else
                {
                    SpaceSplittedString = string.Empty;
                    SpaceCombinedString = string.Empty;
                    while (AttributeValueSlNoString.Length + InputStringToSplit.Length > MaxCharacters)
                    {
                        SpaceSplittedString = InputStringToSplit.Substring(InputStringToSplit.LastIndexOf(' '));
                        if (SpaceCombinedString.Length > 0)
                            SpaceCombinedString = SpaceSplittedString + SpaceCombinedString;
                        else
                            SpaceCombinedString = SpaceSplittedString;
                        InputStringToSplit = InputStringToSplit.Trim().Substring(0, InputStringToSplit.LastIndexOf(' '));
                    }
                    SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit.Trim());
                    InputStringToSplit = string.Empty;
                    AttributeValueSlNo++;
                    if (AttributeValueSlNo.ToString().Length == 1)
                        AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                    else if (AttributeValueSlNo.ToString().Length == 2)
                        AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                    AttributeValueSlNoString += ")";
                    if (SpaceCombinedString.Trim().Length > 0)
                        SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                    //SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit);
                    return SplittedStringList;
                }
            }

            while (true)
            {
                if (InputStringToSplit.IndexOf(',') > 0)
                {
                    if (AttributeValueSlNoString.Length + MaxCharactersMeaningfulString.Length + InputStringToSplit.IndexOf(',') + 1 <= MaxCharacters)
                    {
                        if (string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                            MaxCharactersMeaningfulString = InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(','));
                        else
                            MaxCharactersMeaningfulString += "," + InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(','));

                        InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',') + 1).Trim();
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                        {
                            SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                            MaxCharactersMeaningfulString = string.Empty;
                            AttributeValueSlNo++;
                            if (AttributeValueSlNo.ToString().Length == 1)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                            else if (AttributeValueSlNo.ToString().Length == 2)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                            AttributeValueSlNoString += ")";
                        }
                        else
                        {
                            if (InputStringToSplit.IndexOf(',') > 0)
                            {
                                MaxCharactersMeaningfulString = InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(','));
                                SpaceSplittedString = string.Empty;
                                SpaceCombinedString = string.Empty;
                                while (AttributeValueSlNoString.Length + MaxCharactersMeaningfulString.Length > MaxCharacters)
                                {
                                    SpaceSplittedString = MaxCharactersMeaningfulString.Substring(MaxCharactersMeaningfulString.LastIndexOf(' '));
                                    if (SpaceCombinedString.Length > 0)
                                        SpaceCombinedString = SpaceSplittedString + SpaceCombinedString;
                                    else
                                        SpaceCombinedString = SpaceSplittedString;
                                    MaxCharactersMeaningfulString = MaxCharactersMeaningfulString.Trim().Substring(0, MaxCharactersMeaningfulString.LastIndexOf(' '));
                                }

                                if (SpaceCombinedString.Trim().Length > 0)
                                    InputStringToSplit = SpaceCombinedString.Trim() + "," + InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',') + 1).Trim();
                                else
                                    InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',') + 1).Trim();
                            }
                            else
                            {
                                MaxCharactersMeaningfulString = InputStringToSplit;
                                InputStringToSplit = string.Empty;
                                if (!string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                                {
                                    SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                                    MaxCharactersMeaningfulString = string.Empty;
                                    AttributeValueSlNo++;
                                    if (AttributeValueSlNo.ToString().Length == 1)
                                        AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                                    else if (AttributeValueSlNo.ToString().Length == 2)
                                        AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                                    AttributeValueSlNoString += ")";
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (AttributeValueSlNoString.Length + MaxCharactersMeaningfulString.Length + InputStringToSplit.Length + 1 <= MaxCharacters)
                    {
                        MaxCharactersMeaningfulString += "," + InputStringToSplit;
                        InputStringToSplit = string.Empty;
                        SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                        MaxCharactersMeaningfulString = string.Empty;
                        AttributeValueSlNo++;
                        if (AttributeValueSlNo.ToString().Length == 1)
                            AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                        else if (AttributeValueSlNo.ToString().Length == 2)
                            AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                        AttributeValueSlNoString += ")";
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                        {
                            SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                            MaxCharactersMeaningfulString = string.Empty;
                            AttributeValueSlNo++;
                            if (AttributeValueSlNo.ToString().Length == 1)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                            else if (AttributeValueSlNo.ToString().Length == 2)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                            AttributeValueSlNoString += ")";
                        }

                        if (!string.IsNullOrEmpty(InputStringToSplit))
                        {
                            SpaceSplittedString = string.Empty;
                            SpaceCombinedString = string.Empty;
                            while (AttributeValueSlNoString.Length + InputStringToSplit.Length > MaxCharacters)
                            {
                                SpaceSplittedString = InputStringToSplit.Substring(InputStringToSplit.LastIndexOf(' '));
                                if (SpaceCombinedString.Length > 0)
                                    SpaceCombinedString = SpaceSplittedString + SpaceCombinedString;
                                else
                                    SpaceCombinedString = SpaceSplittedString;
                                InputStringToSplit = InputStringToSplit.Trim().Substring(0, InputStringToSplit.LastIndexOf(' '));
                            }
                            SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit.Trim());
                            InputStringToSplit = string.Empty;
                            AttributeValueSlNo++;
                            if (AttributeValueSlNo.ToString().Length == 1)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                            else if (AttributeValueSlNo.ToString().Length == 2)
                                AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(1, '0');
                            AttributeValueSlNoString += ")";
                            if (SpaceCombinedString.Trim().Length > 0)
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                        }
                    }
                }


                if (string.IsNullOrEmpty(InputStringToSplit))
                    break;

            }

            return SplittedStringList;
        }
    }
}
