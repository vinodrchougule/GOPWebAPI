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
using System.Threading;
using System.Web;
using System.Web.DynamicData;
using System.Web.Http;
using System.Web.Http.Results;

namespace GOPWebAPI.Controllers.GAT
{

    [RoutePrefix("api/DuplicatesCheckOnMultipleColumns")]
    public class DuplicatesCheckOnMultipleColumnsController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            if (!File.Exists(UploadedFilepath))
                return Request.CreateResponse(HttpStatusCode.BadRequest, "File not found");
            else
            {
                string extension = Path.GetExtension(UploadedFilepath);
                if (extension != ".xlsx")
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Invalid file format. Please upload an Excel file.");
                else
                {
                    try
                    {
                        Workbook workbook = new Workbook();
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        workbook.Open(UploadedFilepath);
                        Worksheet ws1 = workbook.Worksheets[0];
                        if (ws1.Cells.Rows.Count <= 1)
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
            }
        }
        #endregion

        #region Check for Duplicate Rows and Write the duplicate remarks in Output to file
        [HttpPost]
        [Route("CheckDuplicatesOnMultipleColumnsAndWriteDuplicateRows")]
        public HttpResponseMessage CheckDuplicatesOnMultipleColumnsAndWriteDuplicateRows(string InputFileName, string UploadedInputFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);
                DataFormatConverter dataFormatConverter = new DataFormatConverter();

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);
                Worksheet wsIF1 = wbIF.Worksheets[0];

                int IFMaxRows1 = wsIF1.Cells.MaxRow;
                int IFMaxColumns1 = wsIF1.Cells.MaxColumn;

                #region Write headings in output worksheet from first worksheet
                if (wbIF.Worksheets.Count == 1)
                    wbIF.Worksheets.Add();

                Worksheet wsIF2 = wbIF.Worksheets[1];
                wsIF2.Name = "Output";

                for (int col = 0; col <= IFMaxColumns1; col++)
                {
                    wsIF2.Cells[0, col].PutValue(wsIF1.Cells[0, col].StringValue);
                    wsIF2.Cells[0, col].SetStyle(wsIF1.Cells[0, col].GetStyle());
                }

                int IFMaxColumns2 = wsIF2.Cells.MaxColumn;
                int IFMaxRow2 = wsIF2.Cells.MaxRow;
                int InputMatchedColNo = IFMaxColumns2 + 1;
                wsIF2.Cells[0, InputMatchedColNo].PutValue("Input Matched");
                int InputMatchedColumnNameColNo = wsIF2.Cells.MaxColumn + 1;
                wsIF2.Cells[0, InputMatchedColumnNameColNo].PutValue("Input Matched Column");
                int PartNoGroupColNo = wsIF2.Cells.MaxColumn + 1;
                wsIF2.Cells[0, PartNoGroupColNo].PutValue("Part No. Group");
                int PartNoGroupCounter = 0;
                int DuplicateSetNoColNo = wsIF2.Cells.MaxColumn + 1;
                wsIF2.Cells[0, DuplicateSetNoColNo].PutValue("Duplicate Set No.");
                #endregion

                string ValueToSearch = string.Empty, ValueToSearchNormalized = string.Empty, cellValue = string.Empty, cellValueNormalized = string.Empty;
                string[] arrCellValueSplittedOnSpace;

                //TODO: Get the ID column as an input, here it is assumed that first column from input file first worksheet is ID column
                //TODO: Get all MFR Part No. columns as an input, here it is assumed that Column No. 2 is MFR Part No1 (Input), Column No. 4 is MFR Part No2, Column No. 6 is MFR Part No3

                #region Adding Combined string column in input sheet
                int CombinedStringColumnNo = wsIF1.Cells.MaxColumn + 1;
                wsIF1.Cells[0, CombinedStringColumnNo].PutValue("Combined String Normalized");
                #endregion

                #region Writing combined string in input sheet
                string CombinedString = string.Empty;
                for (int Row = 1; Row <= IFMaxRows1; Row++)
                {
                    CombinedString = string.Empty;
                    for (int Col = 1; Col <= IFMaxColumns1; Col++)
                    {
                        CombinedString += wsIF1.Cells[Row, Col].StringValue.Trim();
                    }
                    wsIF1.Cells[Row, CombinedStringColumnNo].PutValue(dataFormatConverter.RemoveSpecialCharacters(CombinedString));
                }
                #endregion

                #region Create a list of unique part nos. MFR Part No1(Col.No.2), MFR Part No2(Col.No.4), MFR Part No3(Col.No.6)
                string PartNo = string.Empty;
                List<string> UniquePartNoList = new List<string>();
                for (int Row = 1; Row <= IFMaxRows1; Row++)
                {
                    //MFR Part No1 or Input
                    PartNo = dataFormatConverter.RemoveSpecialCharacters(wsIF1.Cells[Row, 2].StringValue.Trim());
                    if (PartNo.Length > 2 || PartNo.Any(x => char.IsLetter(x)))
                    {
                        if (!UniquePartNoList.Contains(PartNo))
                            UniquePartNoList.Add(PartNo);
                    }

                    //MFR Part No2
                    PartNo = dataFormatConverter.RemoveSpecialCharacters(wsIF1.Cells[Row, 4].StringValue.Trim());
                    if (PartNo.Length > 2 || PartNo.Any(x => char.IsLetter(x)))
                    {
                        if (!UniquePartNoList.Contains(PartNo))
                            UniquePartNoList.Add(PartNo);
                    }

                    //MFR Part No3
                    PartNo = dataFormatConverter.RemoveSpecialCharacters(wsIF1.Cells[Row, 6].StringValue.Trim());
                    if (PartNo.Length > 2 || PartNo.Any(x => char.IsLetter(x)))
                    {
                        if (!UniquePartNoList.Contains(PartNo))
                            UniquePartNoList.Add(PartNo);
                    }
                }
                #endregion

                UniquePartNoList = UniquePartNoList.Distinct().ToList();

                int MatchedRowNo = -1, MatchedColumnNo = -1;
                int sRow = 1, foundRow = -1;
                string MaterialCode = string.Empty, DuplicateSetNo = string.Empty;
                int PartNoGroup = 0, DuplicateSetCounter = 1;
                List<MatchedRowCol> MatchedRowColList = new List<MatchedRowCol>();

                #region Navigate through list and get matching rows and write to output worksheet
                for (int idx = 0; idx < UniquePartNoList.Count(); idx++)
                {
                    ValueToSearch = UniquePartNoList[idx];
                    ValueToSearchNormalized = dataFormatConverter.RemoveSpecialCharacters(UniquePartNoList[idx]);

                    #region Search for normalized value in normalized combined string column
                    MatchedRowColList.Clear();
                    sRow = 1; foundRow = 0;
                    while (sRow <= IFMaxRows1)
                    {
                        if (wsIF1.Cells[sRow, CombinedStringColumnNo].StringValue.ToUpper().Contains(ValueToSearchNormalized.ToUpper()))
                            foundRow = sRow;
                        else
                            foundRow = 0;

                        if (foundRow > 0)
                        {
                            #region Goto left and get the column no.
                            MatchedColumnNo = -1;
                            for (int col = 1; col <= IFMaxColumns1; col++)
                            {
                                cellValue = wsIF1.Cells[foundRow, col].StringValue.Trim();
                                cellValueNormalized = dataFormatConverter.RemoveSpecialCharacters(wsIF1.Cells[foundRow, col].StringValue.Trim());
                                if (cellValueNormalized.Length < ValueToSearchNormalized.Length)        //Cell Value length must be greater than or equal to Value To Search
                                    continue;
                                if (col == 2 || col == 4 || col == 6)    //MFR Part No1, MFR Part No2, MFR Part No3 for exact match
                                {
                                    if (ValueToSearchNormalized.ToUpper() == cellValueNormalized.ToUpper())
                                    {
                                        MatchedColumnNo = col;
                                        break;
                                    }
                                }
                                else
                                {
                                    #region split cell value based on space
                                    arrCellValueSplittedOnSpace = cellValue.Split(' ');
                                    for (int i = 0; i < arrCellValueSplittedOnSpace.Count(); i++)
                                    {
                                        if (dataFormatConverter.RemoveSpecialCharacters(arrCellValueSplittedOnSpace[i].Trim().ToUpper()) == ValueToSearchNormalized.ToUpper())   //Normalized splitted Value Matches to Normalized Value To Search
                                        {
                                            MatchedColumnNo = col;
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                            }
                            #endregion

                            #region if matched cell found, then add it to list
                            if (MatchedColumnNo > -1)
                            {
                                MatchedRowColList.Add(new MatchedRowCol()
                                {
                                    MatchedRowNo = foundRow,
                                    MatchedColNo = MatchedColumnNo
                                });
                            }
                            #endregion
                        }
                        sRow++;
                    }
                    #endregion

                    #region If there are multiple rows matching to input part then only write all the rows to Output worksheet
                    if (MatchedRowColList.Count() > 1)
                    {
                        PartNoGroupCounter++;
                        for (int midx = 0; midx < MatchedRowColList.Count(); midx++)
                        {
                            MatchedRowNo = MatchedRowColList[midx].MatchedRowNo;
                            MatchedColumnNo = MatchedRowColList[midx].MatchedColNo;
                            IFMaxRow2 = wsIF2.Cells.MaxRow + 1;
                            for (int c = 0; c <= IFMaxColumns1; c++)                     //Till last input column
                            {
                                wsIF2.Cells[IFMaxRow2, c].PutValue(wsIF1.Cells[MatchedRowNo, c].StringValue);
                                wsIF2.Cells[IFMaxRow2, c].SetStyle(wsIF1.Cells[MatchedRowNo, c].GetStyle());
                            }
                            wsIF2.Cells[IFMaxRow2, InputMatchedColNo].PutValue(wsIF1.Cells[MatchedRowNo, MatchedColumnNo].StringValue.Trim());
                            wsIF2.Cells[IFMaxRow2, InputMatchedColumnNameColNo].PutValue(wsIF1.Cells[0, MatchedColumnNo].StringValue.Trim());
                            wsIF2.Cells[IFMaxRow2, PartNoGroupColNo].PutValue(PartNoGroupCounter);

                        }
                    }
                    #endregion
                }
                #endregion

                #region Writing duplicate set no.
                List<int> MaterialsPartNoGroupList = new List<int>();
                for (int mRow = 1; mRow <= wsIF2.Cells.MaxRow; mRow++)
                {
                    if (string.IsNullOrEmpty(wsIF2.Cells[mRow, DuplicateSetNoColNo].StringValue.Trim()))
                    {
                        DuplicateSetNo = "A" + DuplicateSetCounter.ToString();
                        wsIF2.Cells[mRow, DuplicateSetNoColNo].PutValue(DuplicateSetNo);
                        MaterialCode = wsIF2.Cells[mRow, 0].StringValue.Trim();
                        PartNoGroup = Convert.ToInt32(wsIF2.Cells[mRow, PartNoGroupColNo].StringValue.Trim());
                        DuplicateSetCounter++;
                    }
                    else
                    {
                        DuplicateSetNo = wsIF2.Cells[mRow, DuplicateSetNoColNo].StringValue.Trim();
                        MaterialCode = wsIF2.Cells[mRow, 0].StringValue.Trim();
                        PartNoGroup = Convert.ToInt32(wsIF2.Cells[mRow, PartNoGroupColNo].StringValue.Trim());
                    }

                    MaterialsPartNoGroupList.Clear();

                    for (int iRow = 1; iRow <= wsIF2.Cells.MaxRow; iRow++)
                    {
                        if (wsIF2.Cells[iRow, 0].StringValue.Trim().ToUpper() == MaterialCode.Trim().ToUpper())
                        {
                            MaterialsPartNoGroupList.Add(Convert.ToInt32(wsIF2.Cells[iRow, PartNoGroupColNo].StringValue.Trim()));
                            if (string.IsNullOrEmpty(wsIF2.Cells[iRow, DuplicateSetNoColNo].StringValue.Trim()))
                                wsIF2.Cells[iRow, DuplicateSetNoColNo].PutValue(DuplicateSetNo);
                        }
                    }

                    MaterialsPartNoGroupList = MaterialsPartNoGroupList.Distinct().ToList();
                    for (int jRow = 1; jRow <= wsIF2.Cells.MaxRow; jRow++)           //mRow + 1
                    {
                        if (string.IsNullOrEmpty(wsIF2.Cells[jRow, DuplicateSetNoColNo].StringValue.Trim()))
                        {
                            foreach (int pn in MaterialsPartNoGroupList)
                            {
                                if (Convert.ToInt32(wsIF2.Cells[jRow, PartNoGroupColNo].StringValue.Trim()) == pn)
                                    wsIF2.Cells[jRow, DuplicateSetNoColNo].PutValue(DuplicateSetNo);
                            }
                        }
                    }
                }
                #endregion

                #region Save and Download the file
                wsIF1.AutoFitColumns();
                wsIF2.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName.Trim();
                wbIF.Save(filename);
                #endregion

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

    public class MatchedRowCol
    {
        public int MatchedRowNo { get; set; }
        public int MatchedColNo { get; set; }
    }
}
