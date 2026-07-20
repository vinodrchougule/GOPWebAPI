using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/FuzzyValueFixer")]
    public class FuzzyValueFixerController : ApiController
    {
        #region Validate the input file
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

                if(wbIF.Worksheets.Count < 2)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File should have at least 2 worksheets");
                }

                #region Validating the first input worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Sl No' not found in input file first worksheet. Please select a valid file.");
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

                if (ws1.Cells.MaxRow <= 0)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first Worksheet has no data rows.");
                }
                #endregion

                #region Validating the second input worksheet
                Worksheet ws2 = wbIF.Worksheets[1];

                if (ws2.Cells[0, 0].StringValue.Trim().ToUpper() != "EXISTING VALUES")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Existing Values' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 1].StringValue.Trim().ToUpper() != "VALUE TYPE-1")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Value Type-1' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 2].StringValue.Trim().ToUpper() != "VALUE TYPE-2")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Value Type-2' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 3].StringValue.Trim().ToUpper() != "SIMILARITY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Similarity' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 4].StringValue.Trim().ToUpper() != "DECISION TO RETAIN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Decision To Retain' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 5].StringValue.Trim().ToUpper() != "REPLACE WITH")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'Replace With' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells.MaxRow <= 0)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file second Worksheet has no data rows.");
                }

                for (int row = 1; row <= ws2.Cells.MaxRow; row++)
                {
                    if (ws2.Cells[row, 4].StringValue.Trim().ToUpper() != "TYPE-1" &&
                        ws2.Cells[row, 4].StringValue.Trim().ToUpper() != "TYPE-2" &&
                        (ws2.Cells[row, 4].StringValue.Trim().ToUpper() != ""))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "In second worksheet, Value in 'Decision To Retain' column has to be either 'Type-1' or 'Type-2' or it should be empty.");
                    }

                    if (string.IsNullOrEmpty(ws2.Cells[row, 4].StringValue.Trim()) && string.IsNullOrEmpty(ws2.Cells[row, 5].StringValue.Trim()))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "In second worksheet, Value in 'Decision To Retain' and Value in 'Replace With' columns both cannot be empty.");
                    }
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully");
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

        #region Fix the fuzzy values and write it to output file
        [HttpPost]
        [Route("FixFuzzyValuesAndWriteToOutput")]
        public HttpResponseMessage FixFuzzyValuesAndWriteToOutput(string UploadedInputFileName, string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                string ValueType1 = string.Empty, ValueType2 = string.Empty, DecisionToRetain = string.Empty, ReplaceWith = string.Empty, ObservationComment = string.Empty;

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet ws1 = wbIF.Worksheets[0];
                Worksheet ws2 = wbIF.Worksheets[1];
                Aspose.Cells.Style styleLightRedFillColor = ws1.Cells[1, 7].GetStyle();

                int RepeatedRowsSetNo = 1, ConflictingDecisionSetNo = 1;
                ws2.Cells[0, 6].PutValue("Observation");

                #region Checking repeated rows
                for (int cRow = 1; cRow <= ws2.Cells.MaxRow; cRow++)
                {
                    ValueType1 = ws2.Cells[cRow, 1].StringValue.Trim().ToUpper();
                    ValueType2 = ws2.Cells[cRow, 2].StringValue.Trim().ToUpper();
                    for (int iRow = cRow + 1; iRow <= ws2.Cells.MaxRow; iRow++)
                    {
                        if (ws2.Cells[iRow, 1].StringValue.Trim().ToUpper() == ValueType1 && ws2.Cells[iRow, 2].StringValue.Trim().ToUpper() == ValueType2)
                        {
                            ws2.Cells[cRow, 6].PutValue("Repeated Values - " + RepeatedRowsSetNo.ToString());
                            ws2.Cells[iRow, 6].PutValue("Repeated Values - " + RepeatedRowsSetNo.ToString());
                            RepeatedRowsSetNo++;
                        }
                    }
                }
                #endregion

                #region Checking conflicting values of rows
                for (int cRow = 1; cRow <= ws2.Cells.MaxRow; cRow++)
                {
                    if (string.IsNullOrEmpty(ws2.Cells[cRow, 5].StringValue.Trim().ToUpper()))
                    {
                        ValueType1 = ws2.Cells[cRow, 1].StringValue.Trim().ToUpper();
                        ValueType2 = ws2.Cells[cRow, 2].StringValue.Trim().ToUpper();
                        DecisionToRetain = ws2.Cells[cRow, 4].StringValue.Trim().ToUpper();

                        for (int iRow = cRow + 1; iRow <= ws2.Cells.MaxRow; iRow++)
                        {
                            if (ws2.Cells[iRow, 2].StringValue.Trim().ToUpper() == ValueType1)
                            {
                                if (ws2.Cells[iRow, 1].StringValue.Trim().ToUpper() == ValueType2)
                                {
                                    if (ws2.Cells[iRow, 4].StringValue.Trim().ToUpper() == DecisionToRetain)
                                    {
                                        ws2.Cells[cRow, 6].PutValue("Conflicting Decision - " + ConflictingDecisionSetNo.ToString());
                                        ws2.Cells[iRow, 6].PutValue("Conflicting Decision - " + ConflictingDecisionSetNo.ToString());
                                        ConflictingDecisionSetNo++;
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Copying the Value column from first worksheet and renaming heading
                ws1.Cells.CopyColumn(ws1.Cells, ws1.Cells.Columns[5].Index, ws1.Cells.Columns[6].Index);
                ws1.Cells[0, 6].PutValue("Value with Replacement");
                #endregion

                #region Reading the second worksheet
                for (int w2Row = 1; w2Row <= ws2.Cells.MaxRow; w2Row++)
                {
                    ValueType1 = ws2.Cells[w2Row, 1].StringValue.Trim().ToUpper();
                    ValueType2 = ws2.Cells[w2Row, 2].StringValue.Trim().ToUpper();
                    DecisionToRetain = ws2.Cells[w2Row, 4].StringValue.Trim().ToUpper();
                    ReplaceWith = ws2.Cells[w2Row, 5].StringValue.Trim().ToUpper();
                    ObservationComment = ws2.Cells[w2Row, 6].StringValue.Trim().ToUpper();

                    if (string.IsNullOrEmpty(ObservationComment))
                    {
                        if (!string.IsNullOrEmpty(ReplaceWith))
                        {
                            for (int w1Row = 1; w1Row <= ws1.Cells.MaxRow; w1Row++)
                            {
                                if (ws1.Cells[w1Row, 6].StringValue.Trim().ToUpper() == ValueType1 || ws1.Cells[w1Row, 6].StringValue.Trim().ToUpper() == ValueType2)
                                {
                                    ws1.Cells[w1Row, 6].PutValue(ReplaceWith);
                                    styleLightRedFillColor = ws1.Cells[w1Row, 6].GetStyle();
                                    styleLightRedFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                    styleLightRedFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                    styleLightRedFillColor.Pattern = BackgroundType.VerticalStripe;
                                    styleLightRedFillColor.Font.Color = System.Drawing.Color.Black;
                                    ws1.Cells[w1Row, 6].SetStyle(styleLightRedFillColor);
                                }
                            }
                        }
                        else
                        {
                            if (DecisionToRetain == "TYPE-1")
                            {
                                for (int w1Row = 1; w1Row <= ws1.Cells.MaxRow; w1Row++)
                                {
                                    if (ws1.Cells[w1Row, 6].StringValue.Trim().ToUpper() == ValueType2)
                                    {
                                        ws1.Cells[w1Row, 6].PutValue(ws2.Cells[w2Row, 1].StringValue);
                                        styleLightRedFillColor = ws1.Cells[w1Row, 6].GetStyle();
                                        styleLightRedFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                        styleLightRedFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                        styleLightRedFillColor.Pattern = BackgroundType.VerticalStripe;
                                        styleLightRedFillColor.Font.Color = System.Drawing.Color.Black;
                                        ws1.Cells[w1Row, 6].SetStyle(styleLightRedFillColor);
                                    }
                                }
                            }
                            else if (DecisionToRetain == "TYPE-2")
                            {
                                for (int w1Row = 1; w1Row <= ws1.Cells.MaxRow; w1Row++)
                                {
                                    if (ws1.Cells[w1Row, 6].StringValue.Trim().ToUpper() == ValueType1)
                                    {
                                        ws1.Cells[w1Row, 6].PutValue(ws2.Cells[w2Row, 2].StringValue);
                                        styleLightRedFillColor = ws1.Cells[w1Row, 6].GetStyle();
                                        styleLightRedFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                        styleLightRedFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                                        styleLightRedFillColor.Pattern = BackgroundType.VerticalStripe;
                                        styleLightRedFillColor.Font.Color = System.Drawing.Color.Black;
                                        ws1.Cells[w1Row, 6].SetStyle(styleLightRedFillColor);
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Save the file
                ws1.AutoFitColumns();
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

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
    }
}
