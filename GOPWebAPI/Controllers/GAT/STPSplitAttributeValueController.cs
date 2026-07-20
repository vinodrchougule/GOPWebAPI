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
    [RoutePrefix("api/STPSplitAttributeValue")]
    public class STPSplitAttributeValueController : ApiController
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'sl no' not found in input file first worksheet. Please select a valid file.");
                }
                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "NM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'NM' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "CLASS_NAME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Class_Name' not found in input file first worksheet. Please select a valid file.");
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

                if (ws1.Cells[0, 7].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [H1] with Value 'Attribute' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 8].StringValue.Trim().ToUpper() != "VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [I1] with Value 'Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                for (int c = 0; c <= 9; c++)
                {
                    if (ws2.Cells[0, c].StringValue.Trim().ToUpper() != ws1.Cells[0, c].StringValue.Trim().ToUpper())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [1," + (c + 1).ToString() + "] with Value '" + ws1.Cells[0, c].StringValue.Trim() + "' not found in input file second worksheet. Please select a valid file.");
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
        [Route("SplitAttributeValueAndWriteToOutput")]
        public HttpResponseMessage SplitAttributeValueAndWriteToOutput(string UploadedInputFileName, string InputFileName, int MaxChars)
        {
            int ir = 1;
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet iws = wbIF.Worksheets[0];
                Worksheet ows = wbIF.Worksheets[1];

                #region Processing the rows from input file first worksheet
                int inputMaxRows = iws.Cells.MaxRow;
                int LastRowFromOutputWroksheet = 1;
                string AttributeValue = string.Empty, AdditionalFeaturesEnSequenceString = string.Empty;
                List<string> MeaningfulAttributeValueSplittedString = new List<string>();
                for (ir = 1; ir <= inputMaxRows; ir++)     //Processing from 1st Row to Last Row of Input Worksheet
                {
                    AttributeValue = iws.Cells[ir, 8].StringValue.Trim();
                    if (AttributeValue.Length > MaxChars && (AttributeValue.IndexOf(',') >= 0 || AttributeValue.IndexOf(' ') >= 0))
                    {
                        if (iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en" || iws.Cells[ir, 7].StringValue.Trim().ToLower() == "equipment type")
                        {
                            if (AttributeValue.ToLower().StartsWith("01)"))
                                AttributeValue = AttributeValue.Substring(3);
                            else if (AttributeValue.ToLower().StartsWith("en_01)"))
                                AttributeValue = AttributeValue.Substring(6);

                            if (iws.Cells[ir, 7].StringValue.Trim().ToLower() == "additional features en")
                                MeaningfulAttributeValueSplittedString = SplitAttributeValueString(AttributeValue, true, MaxChars);
                            else
                                MeaningfulAttributeValueSplittedString = SplitAttributeValueString(AttributeValue, false, MaxChars);
                        }
                        else
                        {
                            if (AttributeValue.ToLower().StartsWith("en_01)"))
                                AttributeValue = AttributeValue.Substring(6);
                            MeaningfulAttributeValueSplittedString = SplitAttributeValueString(AttributeValue, false, MaxChars);
                        }

                        foreach (string ms in MeaningfulAttributeValueSplittedString)
                        {
                            #region Write splitted meaningfule strings to the output worksheet
                            LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                            ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(iws.Cells[ir, 0].StringValue);     //SL No
                            ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(iws.Cells[ir, 1].StringValue);     //Reference
                            ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(iws.Cells[ir, 2].StringValue);     //NM
                            ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue);     //CLASS_NAME
                            ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue);     //Noun
                            ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue);     //Modifier
                            ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(iws.Cells[ir, 6].StringValue);     //Seq
                            ows.Cells[LastRowFromOutputWroksheet + 1, 7].PutValue(iws.Cells[ir, 7].StringValue);     //Attribute
                            ows.Cells[LastRowFromOutputWroksheet + 1, 8].PutValue(ms);                               //Value
                            #endregion
                        }
                    }
                    else
                    {
                        #region write the input row to output as is
                        LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                        ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(iws.Cells[ir, 0].StringValue.Trim());     //SL No
                        ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(iws.Cells[ir, 1].StringValue.Trim());     //Reference
                        ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(iws.Cells[ir, 2].StringValue.Trim());     //NM
                        ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue.Trim());     //CLASS_NAME
                        ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue.Trim());     //Noun
                        ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue.Trim());     //Modifier
                        ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(iws.Cells[ir, 6].StringValue.Trim());     //Seq
                        ows.Cells[LastRowFromOutputWroksheet + 1, 7].PutValue(iws.Cells[ir, 7].StringValue.Trim());     //Attribute
                        ows.Cells[LastRowFromOutputWroksheet + 1, 8].PutValue(iws.Cells[ir, 8].StringValue.Trim());     //Value
                        #endregion
                    }
                }
                #endregion

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

        #region Split Attribute Value String
        private List<string> SplitAttributeValueString(string inputString, bool IsAdditionalFeaturesENAttribute, int MaxCharacters = 70)
        {
            List<string> SplittedStringList = new List<string>();
            string InputStringToSplit = inputString.Trim();
            string MaxCharactersMeaningfulString = string.Empty, AttributeValueSlNoString = string.Empty;
            string SpaceSplittedString = string.Empty, SpaceCombinedString = string.Empty, SpaceCombinedString1 = string.Empty;
            int AttributeValueSlNo = 1;

            if (IsAdditionalFeaturesENAttribute)
                AttributeValueSlNoString = "01)";
            else
                AttributeValueSlNoString = "EN_01)";

            if (InputStringToSplit.IndexOf(',') < 0)
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
                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit);
                if (SpaceCombinedString.Trim().Length > 0)
                {
                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                    SpaceSplittedString = string.Empty;
                    while (AttributeValueSlNoString.Length + SpaceCombinedString.Length > MaxCharacters)
                    {
                        SpaceSplittedString = SpaceCombinedString.Substring(SpaceCombinedString.LastIndexOf(' '));
                        if (SpaceCombinedString1.Length > 0)
                            SpaceCombinedString1 = SpaceSplittedString + SpaceCombinedString1;
                        else
                            SpaceCombinedString1 = SpaceSplittedString;
                        SpaceCombinedString = SpaceCombinedString.Trim().Substring(0, SpaceCombinedString.LastIndexOf(' '));
                    }
                    SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                    SpaceCombinedString = string.Empty;
                    if (SpaceCombinedString1.Trim().Length > 0)
                    {
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                        SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString1.Trim());
                        SpaceCombinedString1 = string.Empty;
                    }
                }

                return SplittedStringList;
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
                            AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
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
                            SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());

                            if (SpaceCombinedString.Trim().Length > 0)
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                                SpaceSplittedString = string.Empty;
                                while (AttributeValueSlNoString.Length + SpaceCombinedString.Length > MaxCharacters)
                                {
                                    SpaceSplittedString = SpaceCombinedString.Substring(SpaceCombinedString.LastIndexOf(' '));
                                    if (SpaceCombinedString1.Length > 0)
                                        SpaceCombinedString1 = SpaceSplittedString + SpaceCombinedString1;
                                    else
                                        SpaceCombinedString1 = SpaceSplittedString;
                                    SpaceCombinedString = SpaceCombinedString.Trim().Substring(0, SpaceCombinedString.LastIndexOf(' '));
                                }
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                                SpaceCombinedString = string.Empty;
                                if (SpaceCombinedString1.Trim().Length > 0)
                                {
                                    InputStringToSplit = SpaceCombinedString1 + "," + InputStringToSplit;
                                    //AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                                    //SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString1.Trim());
                                    SpaceCombinedString1 = string.Empty;
                                }
                            }
                            MaxCharactersMeaningfulString = string.Empty;
                        }
                        else
                        {
                            if (InputStringToSplit.IndexOf(',') > 0)
                            {
                                MaxCharactersMeaningfulString = InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(','));
                                InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',') + 1).Trim();
                            }
                            else
                            {
                                MaxCharactersMeaningfulString = InputStringToSplit;
                                InputStringToSplit = string.Empty;
                                if (!string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                                {
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                                    SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                                    MaxCharactersMeaningfulString = string.Empty;
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
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                        SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());
                        MaxCharactersMeaningfulString = string.Empty;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(MaxCharactersMeaningfulString))
                        {
                            AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
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
                            SplittedStringList.Add(AttributeValueSlNoString + MaxCharactersMeaningfulString.Trim());

                            if (SpaceCombinedString.Trim().Length > 0)
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                                SpaceSplittedString = string.Empty;
                                while (AttributeValueSlNoString.Length + SpaceCombinedString.Length > MaxCharacters)
                                {
                                    SpaceSplittedString = SpaceCombinedString.Substring(SpaceCombinedString.LastIndexOf(' '));
                                    if (SpaceCombinedString1.Length > 0)
                                        SpaceCombinedString1 = SpaceSplittedString + SpaceCombinedString1;
                                    else
                                        SpaceCombinedString1 = SpaceSplittedString;
                                    SpaceCombinedString = SpaceCombinedString.Trim().Substring(0, SpaceCombinedString.LastIndexOf(' '));
                                }
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                                SpaceCombinedString = string.Empty;
                                if (SpaceCombinedString1.Trim().Length > 0)
                                {
                                    InputStringToSplit = SpaceCombinedString1 + "," + InputStringToSplit;
                                    SpaceCombinedString1 = string.Empty;
                                }
                            }
                            MaxCharactersMeaningfulString = string.Empty;
                        }

                        if (!string.IsNullOrEmpty(InputStringToSplit))
                        {
                            AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
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

                            if (SpaceCombinedString.Trim().Length > 0)
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, IsAdditionalFeaturesENAttribute);
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString.Trim());
                            }
                            InputStringToSplit = string.Empty;
                        }
                    }
                }

                if (string.IsNullOrEmpty(InputStringToSplit))
                    break;
            }

            return SplittedStringList;
        }
        #endregion

        #region Get Attribute Value SlNo string
        private string GetAttributeValueSlNoString(int AttributeValueSlNo, bool IsAdditionalFeaturesENAttribute)
        {
            string AttributeValueSlNoString = string.Empty;

            if (IsAdditionalFeaturesENAttribute)
            {
                if (AttributeValueSlNo.ToString().Length == 1)
                    AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                else
                    AttributeValueSlNoString = AttributeValueSlNo.ToString();

                AttributeValueSlNoString = AttributeValueSlNoString + ")";
            }
            else
            {
                if (AttributeValueSlNo.ToString().Length == 1)
                    AttributeValueSlNoString = AttributeValueSlNo.ToString().PadLeft(2, '0');
                else
                    AttributeValueSlNoString = AttributeValueSlNo.ToString();

                AttributeValueSlNoString = "EN_" + AttributeValueSlNoString + ")";
            }

            return AttributeValueSlNoString;
        }
        #endregion
    }
}
