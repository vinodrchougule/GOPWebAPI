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
    [RoutePrefix("api/CharacterLengthSplitter")]
    public class CharacterLengthSplitterController : ApiController
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

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 4].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 5].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'Attribute' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 6].StringValue.Trim().ToUpper() != "VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [G1] with Value 'Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                if(wbIF.Worksheets.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File should have at least two worksheets.");
                }

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                for (int c = 0; c <= 7; c++)
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
        public HttpResponseMessage SplitAttributeValueAndWriteToOutput(string UploadedInputFileName, string InputFileName, int MaxChars = 30, string Prefix = "")
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            int ir = 1;

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
                string AttributeValue = string.Empty;
                List<string> MeaningfulAttributeValueSplittedString = new List<string>();
                for (ir = 1; ir <= inputMaxRows; ir++)     //Processing from 1st Row to Last Row of Input Worksheet
                {
                    AttributeValue = iws.Cells[ir, 6].StringValue.Trim();
                    if (AttributeValue.Length > MaxChars && (AttributeValue.IndexOf(',') >= 0 || AttributeValue.IndexOf(' ') >= 0))
                    {
                        MeaningfulAttributeValueSplittedString = SplitAttributeValueString(AttributeValue, MaxChars);

                        foreach (string ms in MeaningfulAttributeValueSplittedString)
                        {
                            #region Write splitted meaningfule strings to the output worksheet
                            LastRowFromOutputWroksheet = ows.Cells.MaxRow;
                            ows.Cells[LastRowFromOutputWroksheet + 1, 0].PutValue(iws.Cells[ir, 0].StringValue);     //SL No
                            ows.Cells[LastRowFromOutputWroksheet + 1, 1].PutValue(iws.Cells[ir, 1].StringValue);     //Reference
                            ows.Cells[LastRowFromOutputWroksheet + 1, 2].PutValue(iws.Cells[ir, 2].StringValue);     //NM
                            ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue);     //Noun
                            ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue);     //Modifier
                            ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue);     //Attribute
                            ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(ms);                               //Value
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
                        ows.Cells[LastRowFromOutputWroksheet + 1, 3].PutValue(iws.Cells[ir, 3].StringValue.Trim());     //Noun
                        ows.Cells[LastRowFromOutputWroksheet + 1, 4].PutValue(iws.Cells[ir, 4].StringValue.Trim());     //Modifier
                        ows.Cells[LastRowFromOutputWroksheet + 1, 5].PutValue(iws.Cells[ir, 5].StringValue.Trim());     //Attribute
                        ows.Cells[LastRowFromOutputWroksheet + 1, 6].PutValue(iws.Cells[ir, 6].StringValue.Trim());     //Value
                        #endregion
                    }
                }
                #endregion

                #region if no. of characters are equal to max. characters and ends with comma, then remove ending comma
                for (int or = 1; or <= ows.Cells.MaxRow; or++)
                {
                    if (ows.Cells[or, 6].StringValue.Length == (MaxChars + 1) && ows.Cells[or, 6].StringValue.EndsWith(","))
                        ows.Cells[or, 6].PutValue(ows.Cells[or, 6].StringValue.Substring(0, MaxChars));
                }
                #endregion

                #region Save the file
                ows.AutoFitColumns();
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
        private List<string> SplitAttributeValueString(string inputString, int MaxCharacters = 40, string Prefix = "")
        {
            List<string> SplittedStringList = new List<string>();
            List<string> tmpSplittedStringList = new List<string>();
            string InputStringToSplit = inputString.Trim(), AttributeValueSlNoString = string.Empty;
            string SpaceSplittedString = string.Empty, SpaceCombinedString = string.Empty, CommaSplittedString = string.Empty, CommaCombinedString = string.Empty;
            int AttributeValueSlNo = 1;

            AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);

            #region Processing Input String not having comma and space both
            if (InputStringToSplit.IndexOf(',') < 0 && InputStringToSplit.IndexOf(' ') < 0)
            {
                SplittedStringList.Add(InputStringToSplit);
                return SplittedStringList;
            }
            #endregion

            #region Processing Input String not having comma and having space
            if (InputStringToSplit.IndexOf(',') < 0)
                return ProcessInputStringNotHavingCommaAndHavingSpace(InputStringToSplit, MaxCharacters, AttributeValueSlNo, Prefix);
            #endregion

            #region Processing Input String having comma and no space
            if (InputStringToSplit.IndexOf(',') >= 0 && InputStringToSplit.IndexOf(' ') < 0)
            {
                if (AttributeValueSlNoString.Length + InputStringToSplit.Length > MaxCharacters)
                {
                    while (true)
                    {
                        if (InputStringToSplit.IndexOf(',') > 0)
                        {
                            CommaSplittedString = InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(','));
                            if (CommaCombinedString.Length + CommaSplittedString.Length + 1 + AttributeValueSlNoString.Length > MaxCharacters)        //1 for comma
                            {
                                if (CommaCombinedString.Trim().Length > 0)
                                {
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                    SplittedStringList.Add(AttributeValueSlNoString + CommaCombinedString);
                                    CommaCombinedString = string.Empty;
                                }

                                if ((CommaSplittedString.Length - 1) + AttributeValueSlNoString.Length >= MaxCharacters)
                                {
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                    SplittedStringList.Add(AttributeValueSlNoString + CommaSplittedString.Trim());

                                    InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',')).Trim();
                                }
                            }
                            else
                            {
                                if (CommaCombinedString.Length > 0)
                                    CommaCombinedString = CommaCombinedString + "," + CommaSplittedString;
                                else
                                    CommaCombinedString = CommaSplittedString;

                                InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(',') + 1).Trim();
                            }
                        }
                        else
                        {
                            if (AttributeValueSlNo > 0)
                            {
                                if (CommaCombinedString.Length + InputStringToSplit.Length + 1 + AttributeValueSlNoString.Length < MaxCharacters)        //1 for comma
                                {
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                    SplittedStringList.Add(AttributeValueSlNoString + CommaCombinedString + "," + InputStringToSplit);
                                }
                                else
                                {
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                    SplittedStringList.Add(AttributeValueSlNoString + CommaCombinedString);
                                    AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                    SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit);
                                }
                            }
                            else
                                SplittedStringList.Add(InputStringToSplit);
                            break;
                        }
                    }
                }
                else
                    SplittedStringList.Add(InputStringToSplit);

                return SplittedStringList;
            }
            #endregion

            #region Processing Input String having comma and space both
            if (InputStringToSplit.IndexOf(',') >= 0 && InputStringToSplit.IndexOf(' ') >= 0)
            {
                SplittedStringList = ProcessInputStringHavingCommaAndSpaceBoth(InputStringToSplit, MaxCharacters, AttributeValueSlNo, Prefix);
            }
            #endregion

            return SplittedStringList;
        }
        #endregion

        #region Get Attribute Value SlNo string
        private string GetAttributeValueSlNoString(int AttributeValueSlNo, string Prefix)
        {
            if (Prefix.Length > 0)
                return Prefix + AttributeValueSlNo.ToString() + ":";
            else
                return "";
        }
        #endregion

        #region Process Input String not having comma and having space
        private List<string> ProcessInputStringNotHavingCommaAndHavingSpace(string InputStringToSplit, int MaxCharacters, int AttributeValueSlNo = 1, string Prefix = "")
        {
            List<string> SplittedStringList = new List<string>();
            string AttributeValueSlNoString = string.Empty;
            string SpaceSplittedString = string.Empty, SpaceCombinedString = string.Empty, CommaSplittedString = string.Empty, CommaCombinedString = string.Empty;
            AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);

            if (AttributeValueSlNoString.Length + InputStringToSplit.Length > MaxCharacters)
            {
                while (true)
                {
                    if (InputStringToSplit.IndexOf(' ') > 0)
                    {
                        SpaceSplittedString = InputStringToSplit.Substring(0, InputStringToSplit.IndexOf(' '));
                        if (SpaceCombinedString.Length + SpaceSplittedString.Length + 1 + AttributeValueSlNoString.Length > MaxCharacters)        //1 for space
                        {
                            if (SpaceCombinedString.Trim().Length > 0)
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString);
                                SpaceCombinedString = string.Empty;
                            }

                            if ((SpaceSplittedString.Length - 1) + AttributeValueSlNoString.Length >= MaxCharacters)
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceSplittedString.Trim());

                                InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(' ')).Trim();
                            }
                        }
                        else
                        {
                            if (SpaceCombinedString.Length > 0)
                                SpaceCombinedString = SpaceCombinedString + " " + SpaceSplittedString;
                            else
                                SpaceCombinedString = SpaceSplittedString;

                            InputStringToSplit = InputStringToSplit.Trim().Substring(InputStringToSplit.IndexOf(' ')).Trim();
                        }
                    }
                    else
                    {
                        if (AttributeValueSlNo > 0)
                        {
                            if (SpaceCombinedString.Length + InputStringToSplit.Length + 1 + AttributeValueSlNoString.Length < MaxCharacters)        //1 for space
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString + " " + InputStringToSplit);
                            }
                            else
                            {
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                SplittedStringList.Add(AttributeValueSlNoString + SpaceCombinedString);
                                AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                                SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit);
                            }
                        }
                        else
                            SplittedStringList.Add(InputStringToSplit);
                        break;
                    }
                }
            }
            else
                SplittedStringList.Add(InputStringToSplit);

            return SplittedStringList;
        }
        #endregion

        #region Process Input String Having Comma And Space Both
        private List<string> ProcessInputStringHavingCommaAndSpaceBoth(string InputStringToSplit, int MaxCharacters, int AttributeValueSlNo = 1, string Prefix = "")
        {
            List<string> SplittedStringList = new List<string>();
            int indexOfSpace = -1, indexOfComma = -1, lesserIndex = -1;
            string SplittedString = string.Empty, CombinedString = string.Empty;
            string AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo, Prefix);

            while (true)
            {
                indexOfSpace = InputStringToSplit.IndexOf(' ');
                indexOfComma = InputStringToSplit.IndexOf(',');

                if (indexOfSpace < 0 && indexOfComma < 0)
                {
                    if (CombinedString.Length + InputStringToSplit.Length + AttributeValueSlNoString.Length > MaxCharacters)
                    {
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                        SplittedStringList.Add(AttributeValueSlNoString + CombinedString.Trim());
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                        SplittedStringList.Add(AttributeValueSlNoString + InputStringToSplit.Trim());
                    }
                    else
                    {
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                        SplittedStringList.Add(AttributeValueSlNoString + CombinedString + InputStringToSplit.Trim());
                    }
                    break;
                }
                else if ((indexOfSpace < indexOfComma || indexOfComma < 0) && indexOfSpace > -1)
                    lesserIndex = indexOfSpace;
                else
                    lesserIndex = indexOfComma;

                SplittedString = InputStringToSplit.Substring(0, lesserIndex + 1);
                if (CombinedString.Length + (SplittedString.Length - 1) + AttributeValueSlNoString.Length > MaxCharacters)
                {
                    if (CombinedString.Trim().Length > 0)
                    {
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                        SplittedStringList.Add(AttributeValueSlNoString + CombinedString.Trim());
                        CombinedString = string.Empty;
                    }

                    if ((SplittedString.Length - 1) + AttributeValueSlNoString.Length >= MaxCharacters)
                    {
                        AttributeValueSlNoString = GetAttributeValueSlNoString(AttributeValueSlNo++, Prefix);
                        SplittedStringList.Add(AttributeValueSlNoString + SplittedString.Trim());

                        InputStringToSplit = InputStringToSplit.Trim().Substring(lesserIndex + 1).Trim();
                    }
                }
                else
                {
                    if (CombinedString.Length > 0)
                        CombinedString = CombinedString + SplittedString;
                    else
                        CombinedString = SplittedString;

                    InputStringToSplit = InputStringToSplit.Trim().Substring(lesserIndex + 1).Trim();
                }
            }

            return SplittedStringList;
        }
        #endregion
    }
}
