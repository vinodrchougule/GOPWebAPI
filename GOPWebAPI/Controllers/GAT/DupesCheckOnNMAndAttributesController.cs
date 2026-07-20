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
using System.Web.Http.Results;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/DupesCheckOnNMAndAttributes")]
    public class DupesCheckOnNMAndAttributesController : ApiController
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
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating first worksheet
                string hdr;
                int attNo = 1;
                for (int aCtr = 2; aCtr <= 100; aCtr+=2)
                {
                    hdr = ws1.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' in input file first worksheet.Please download format from Template.");
                    }

                    hdr = ws1.Cells[0, aCtr + 2].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 3).ToString() + "] must contain 'Attribute Value" + attNo.ToString() + "' in input file first worksheet.Please download format from Template.");
                    }
                    attNo++;
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

        #region Find Duplicates and Write the Output to file
        [HttpPost]
        [Route("FindDuplicatesAndWriteOutputToExcel")]
        public HttpResponseMessage FindDuplicatesAndWriteOutputToExcel(string InputFileName, string UploadedInputFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);

                Worksheet ws1 = wbIF.Worksheets[0];

                int IFMaxRows = ws1.Cells.MaxRow;
                int IFMaxColumns = ws1.Cells.MaxColumn;
                int ResultMaxRows = 0, ResultMaxColumns = 0;

                #region Write headings to result worksheet
                if (wbIF.Worksheets.Count == 1)
                    wbIF.Worksheets.Add("Result");
                else
                    wbIF.Worksheets[1].Name = "Result";

                Worksheet wsResult = wbIF.Worksheets["Result"];

                for (int iCol = 0; iCol <= IFMaxColumns; iCol++)
                    wsResult.Cells[0, iCol].PutValue(ws1.Cells[0, iCol].StringValue.Trim());

                ResultMaxColumns = wsResult.Cells.MaxColumn;
                int NoOfFilledValuesCol = ResultMaxColumns + 1;
                wsResult.Cells[0, NoOfFilledValuesCol].PutValue("No. of Filled Values");
                int NoOfMatchingValuesCol = NoOfFilledValuesCol + 1;
                wsResult.Cells[0, NoOfMatchingValuesCol].PutValue("No. of Matching Values");
                int PercentageOfMatchCol = NoOfMatchingValuesCol + 1;
                wsResult.Cells[0, PercentageOfMatchCol].PutValue("% Match");
                int RemarksCol = PercentageOfMatchCol + 1;
                wsResult.Cells[0, RemarksCol].PutValue("Remarks");
                int AttributeValueFillPercentageCol = RemarksCol + 1;
                wsResult.Cells[0, AttributeValueFillPercentageCol].PutValue("Attribute Value Fill (%)");
                int NoOfAttributesCol = AttributeValueFillPercentageCol + 1;
                wsResult.Cells[0, NoOfAttributesCol].PutValue("No. of Attributes");
                ResultMaxRows = wsResult.Cells.MaxRow;
                #endregion

                int NoOfValuesFilledInMainRow = 0, NoOfValuesFilledInNavigatingRow = 0, NoOfAttributes = 0;
                int NoOfMatchedValues = 0, DuplicateSetCounter = 0;
                string Reference = string.Empty;
                bool IsMainRowWrittenToResultWorksheet = false;
                string cs = string.Empty;
                CellArea ca = new CellArea();
                List<string> UniqueReferenceList = new List<string>();

                #region Writing Combined String
                int CombinedStringCol = ws1.Cells.MaxColumn + 1;
                ws1.Cells[0, CombinedStringCol].PutValue("Combined String");
                for (int iRow = 1; iRow <= IFMaxRows; iRow++)
                {
                    cs = string.Empty;
                    for (int col = 1; col <= IFMaxColumns; col++)
                        cs = cs + ws1.Cells[iRow, col].StringValue.Trim();

                    ws1.Cells[iRow, CombinedStringCol].PutValue(cs);
                }
                #endregion

                #region Processing Input File Rows starts
                Aspose.Cells.Cell cell = ws1.Cells[1, CombinedStringCol];
                for (int mRow = 1; mRow <= IFMaxRows; mRow++)                   //Main Row
                {
                    cs = ws1.Cells[mRow, CombinedStringCol].StringValue.Trim();
                    IsMainRowWrittenToResultWorksheet = false;
                    if (!UniqueReferenceList.Contains(ws1.Cells[mRow, 0].StringValue.Trim()))
                    {
                        ca.StartRow = mRow + 1;
                        ca.StartColumn = CombinedStringCol;
                        ca.EndRow = IFMaxRows;
                        ca.EndColumn = CombinedStringCol;
                        while (true)
                        {
                            cell = ws1.Cells.FindStringContains(cs, cell, false, ca);
                            if (cell == null)
                                break;
                            if (ws1.Cells[mRow, CombinedStringCol].StringValue.Trim().ToUpper() ==
                                ws1.Cells[cell.Row, CombinedStringCol].StringValue.Trim().ToUpper())
                            {
                                if (!UniqueReferenceList.Contains(ws1.Cells[cell.Row, 0].StringValue.Trim()))
                                {
                                    if (!IsMainRowWrittenToResultWorksheet)
                                    {
                                        #region Write main row to result worksheet
                                        DuplicateSetCounter++;

                                        #region Find No. of values filled in main row
                                        NoOfValuesFilledInMainRow = 0;
                                        NoOfAttributes = 0;
                                        for (int mCol = 3; mCol <= IFMaxColumns; mCol += 2)
                                        {
                                            if (string.IsNullOrEmpty(ws1.Cells[mRow, mCol].StringValue.Trim()))        //break if Attribute Name is empty
                                                break;

                                            NoOfAttributes++;
                                            if (!string.IsNullOrEmpty(ws1.Cells[mRow, mCol + 1].StringValue.Trim()))
                                                NoOfValuesFilledInMainRow++;
                                        }
                                        #endregion

                                        if (NoOfValuesFilledInMainRow > 0)
                                        {
                                            #region Writing main row
                                            ResultMaxRows = wsResult.Cells.MaxRow + 1;
                                            wsResult.Cells[ResultMaxRows, 0].PutValue(ws1.Cells[mRow, 0].StringValue.Trim());          //Reference
                                            wsResult.Cells[ResultMaxRows, 1].PutValue(ws1.Cells[mRow, 1].StringValue.Trim());          //Noun
                                            wsResult.Cells[ResultMaxRows, 2].PutValue(ws1.Cells[mRow, 2].StringValue.Trim());          //Modifier
                                            for (int col = 3; col <= IFMaxColumns; col += 2)
                                            {
                                                if (string.IsNullOrEmpty(ws1.Cells[mRow, col].StringValue.Trim()))                     //break if Attribute Name is empty
                                                    break;

                                                wsResult.Cells[ResultMaxRows, col].PutValue(ws1.Cells[mRow, col].StringValue.Trim());      //Attribute Name
                                                wsResult.Cells[ResultMaxRows, col + 1].PutValue(ws1.Cells[mRow, col + 1].StringValue.Trim());  //Attribute Value
                                            }
                                            wsResult.Cells[ResultMaxRows, NoOfFilledValuesCol].PutValue(NoOfValuesFilledInMainRow);
                                            wsResult.Cells[ResultMaxRows, RemarksCol].PutValue("Duplicate Set No." + DuplicateSetCounter.ToString());
                                            if (NoOfAttributes > 0)
                                                wsResult.Cells[ResultMaxRows, AttributeValueFillPercentageCol].PutValue((NoOfValuesFilledInMainRow * 100.00) / NoOfAttributes);
                                            wsResult.Cells[ResultMaxRows, NoOfAttributesCol].PutValue(NoOfAttributes);
                                            #endregion

                                            IsMainRowWrittenToResultWorksheet = true;
                                            UniqueReferenceList.Add(ws1.Cells[mRow, 0].StringValue.Trim());
                                        }
                                        #endregion
                                    }

                                    #region Write navigating row to result worksheet

                                    if (NoOfValuesFilledInMainRow > 0)
                                    {
                                        #region Writing navigating row
                                        NoOfValuesFilledInNavigatingRow = 0;
                                        NoOfAttributes = 0;
                                        NoOfMatchedValues = 0;
                                        ResultMaxRows = wsResult.Cells.MaxRow + 1;
                                        wsResult.Cells[ResultMaxRows, 0].PutValue(ws1.Cells[cell.Row, 0].StringValue.Trim());          //Reference
                                        wsResult.Cells[ResultMaxRows, 1].PutValue(ws1.Cells[cell.Row, 1].StringValue.Trim());          //Noun
                                        wsResult.Cells[ResultMaxRows, 2].PutValue(ws1.Cells[cell.Row, 2].StringValue.Trim());          //Modifier
                                        for (int col = 3; col <= IFMaxColumns; col += 2)
                                        {
                                            if (string.IsNullOrEmpty(ws1.Cells[cell.Row, col].StringValue.Trim()))                     //break if Attribute Name is empty
                                                break;

                                            NoOfAttributes++;
                                            wsResult.Cells[ResultMaxRows, col].PutValue(ws1.Cells[cell.Row, col].StringValue.Trim());          //Attribute Name
                                            if (!string.IsNullOrEmpty(ws1.Cells[cell.Row, col + 1].StringValue.Trim()))
                                            {
                                                wsResult.Cells[ResultMaxRows, col + 1].PutValue(ws1.Cells[cell.Row, col + 1].StringValue.Trim());  //Attribute Value
                                                NoOfValuesFilledInNavigatingRow++;
                                            }

                                            if (!(string.IsNullOrEmpty(ws1.Cells[mRow, col + 1].StringValue.Trim()) &&
                                                  string.IsNullOrEmpty(ws1.Cells[cell.Row, col + 1].StringValue.Trim())))
                                            {
                                                if (ws1.Cells[mRow, col].StringValue.Trim().ToUpper() == ws1.Cells[cell.Row, col].StringValue.Trim().ToUpper() &&      //Attribute Name
                                                    ws1.Cells[mRow, col + 1].StringValue.Trim().ToUpper() == ws1.Cells[cell.Row, col + 1].StringValue.Trim().ToUpper())    //Attribute Value
                                                    NoOfMatchedValues++;
                                            }
                                        }
                                        wsResult.Cells[ResultMaxRows, NoOfFilledValuesCol].PutValue(NoOfValuesFilledInNavigatingRow);
                                        wsResult.Cells[ResultMaxRows, NoOfMatchingValuesCol].PutValue(NoOfMatchedValues);
                                        wsResult.Cells[ResultMaxRows, PercentageOfMatchCol].PutValue(100);
                                        wsResult.Cells[ResultMaxRows, RemarksCol].PutValue("Duplicate Set No." + DuplicateSetCounter.ToString());
                                        if (NoOfAttributes > 0)
                                            wsResult.Cells[ResultMaxRows, AttributeValueFillPercentageCol].PutValue((NoOfValuesFilledInNavigatingRow * 100.00) / NoOfAttributes);
                                        wsResult.Cells[ResultMaxRows, NoOfAttributesCol].PutValue(NoOfAttributes);
                                        #endregion

                                        UniqueReferenceList.Add(ws1.Cells[cell.Row, 0].StringValue.Trim());
                                    }
                                    #endregion
                                }
                            }
                        }
                    }
                }
                #endregion

                wsResult.AutoFitColumns();
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
