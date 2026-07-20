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
    [RoutePrefix("api/DuplicatesCheckGeneric")]
    public class DuplicatesCheckGenericController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

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
                wbIF.Open(UploadedInputFilepath);

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedInputFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }

                #region checking unique column headings
                string ColumnName = "", DuplicateColumnName = "";
                int ColumnHeadingCounter = 0;
                for (int ColNo = 0; ColNo <= ws1.Cells.MaxColumn; ColNo++)
                {
                    ColumnName = ws1.Cells[0, ColNo].StringValue.Trim();
                    for (int c = 0; c <= ws1.Cells.MaxColumn; c++)
                    {
                        if (ColumnName.ToUpper() == ws1.Cells[0, c].StringValue.Trim().ToUpper())
                        {
                            ColumnHeadingCounter++;
                            if (ColumnHeadingCounter > 1)
                            {
                                DuplicateColumnName = ColumnName;
                                break;
                            }
                        }
                    }

                    if (ColumnHeadingCounter > 1)
                    {
                        File.Delete(UploadedInputFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column headings from input first worksheet must be unique. Duplicate column name:'" + DuplicateColumnName + "'");
                    }
                    ColumnHeadingCounter = 0;
                }
                #endregion
                #endregion


                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully.");
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Check for Duplicate Rows and Write the duplicate remarks in Output to file
        [HttpPost]
        [Route("DuplicateCheckGenericAndWriteOutputToExcel")]
        public HttpResponseMessage DuplicateCheckGenericAndWriteOutputToExcel(DuplicatesCheckGenericModel model)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + model.InputFileName);

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Create an array of selected column numbers based on the column names provided in input
                int[] arrSelectedColumnNos;
                List<int> lstSelectedColumnNos = new List<int>();

                foreach(string col in model.ColumnsToCheckForDuplicates)
                {
                    for(int c=0; c<= wsIF.Cells.MaxColumn; c++)
                    {
                        if(col.Trim().ToUpper() == wsIF.Cells[0, c].StringValue.Trim().ToUpper())
                        {
                            lstSelectedColumnNos.Add(c);
                            break;
                        }
                    }
                }
                arrSelectedColumnNos = lstSelectedColumnNos.ToArray();
                #endregion


                string selectedColumnNames = string.Join(",", model.ColumnsToCheckForDuplicates);
                double ApplyPercentageOfMatch = model.PercentageMatch;
                int LastColumnNo = wsIF.Cells.MaxColumn;
                int DuplicatesToWriteInColumnNo = LastColumnNo + 1;
                int UIColumnNo = 0;

                #region Find Unique Identifier column number
                for (int ColNo = 0; ColNo <= wsIF.Cells.MaxColumn; ColNo++)
                {
                    if (wsIF.Cells[0, ColNo].StringValue.Trim().ToUpper() == model.UniqueIdentifierColumnName.Trim().ToUpper())
                    {
                        UIColumnNo = ColNo;
                        break;
                    }
                }
                #endregion

                if (model.DuplicatesToCheckBasedOn == "I")
                {
                    #region Exact duplicates
                    string CellValueToFind = "", CellValueToFindOn = "";
                    int DuplicateSetCounter = 1;
                    int CheckOrOKCommentColumnNo = 0;
                    CellArea ca = new CellArea();
                    bool IsSetPrinted = false;
                    string UnMatchedColumnHeadingOrAttribute = string.Empty;
                    if (model.IsToFindDuplicatesWithExactMatch)
                    {
                        for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                        {
                            Aspose.Cells.Cell cell = wsIF.Cells[1, Convert.ToInt32(arrSelectedColumnNos[index])];
                            DuplicateSetCounter = 1;
                            DuplicatesToWriteInColumnNo = wsIF.Cells.MaxColumn + 1;
                            wsIF.Cells[0, DuplicatesToWriteInColumnNo].PutValue("Exact duplicate on " + wsIF.Cells[0, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim());
                            CheckOrOKCommentColumnNo = wsIF.Cells.MaxColumn + 1;
                            wsIF.Cells[0, CheckOrOKCommentColumnNo].PutValue("Check or OK Comment on " + wsIF.Cells[0, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim());
                            bool AllColumnValuesAreSame = true;
                            for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                            {
                                if (!string.IsNullOrEmpty(wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim()))
                                {
                                    if (string.IsNullOrEmpty(wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim()))
                                    {
                                        CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim().ToUpper();
                                        IsSetPrinted = false;
                                        ca.StartRow = RowNo + 1;
                                        ca.StartColumn = Convert.ToInt32(arrSelectedColumnNos[index]);
                                        ca.EndRow = wsIF.Cells.MaxRow;
                                        ca.EndColumn = Convert.ToInt32(arrSelectedColumnNos[index]);
                                        UnMatchedColumnHeadingOrAttribute = string.Empty;
                                        while (true)
                                        {
                                            cell = wsIF.Cells.FindStringContains(CellValueToFind, cell, false, ca);
                                            if (cell == null)
                                                break;
                                            if (wsIF.Cells[cell.Row, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim().ToUpper() == CellValueToFind.ToUpper())
                                            {
                                                wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter.ToString());
                                                wsIF.Cells[cell.Row, DuplicatesToWriteInColumnNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter.ToString());
                                                IsSetPrinted = true;
                                                AllColumnValuesAreSame = true;
                                                for (int col = 0; col <= LastColumnNo; col++)
                                                {
                                                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() != "additional information" && col != UIColumnNo)
                                                    {
                                                        if (wsIF.Cells[cell.Row, col].StringValue.Trim().ToLower() != wsIF.Cells[RowNo, col].StringValue.Trim().ToLower())
                                                        {
                                                            if (wsIF.Cells[0, col - 1].StringValue.Trim().ToLower().Contains("att") && wsIF.Cells[0, col].StringValue.Trim().ToLower().Contains("value"))
                                                                UnMatchedColumnHeadingOrAttribute = wsIF.Cells[cell.Row, col - 1].StringValue.Trim();
                                                            else
                                                                UnMatchedColumnHeadingOrAttribute = wsIF.Cells[0, col].StringValue.Trim();
                                                            wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("Check");
                                                            wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].PutValue("Check - " + UnMatchedColumnHeadingOrAttribute);
                                                            AllColumnValuesAreSame = false;
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (AllColumnValuesAreSame)
                                                {
                                                    wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("OK");
                                                    wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].PutValue("OK");
                                                }
                                            }
                                        }

                                        if (IsSetPrinted)
                                            DuplicateSetCounter++;
                                    }
                                }
                            }

                            #region Writing Check Comment to all set even if one row has Check Comment
                            string ExactDuplicateSetComment = string.Empty;
                            cell = wsIF.Cells[1, DuplicatesToWriteInColumnNo];
                            for (int RowNo = 1; RowNo <= wsIF.Cells.MaxRow; RowNo++)
                            {
                                if (wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].StringValue.Trim().ToUpper() == "OK")
                                {
                                    ExactDuplicateSetComment = wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim();
                                    ca.StartRow = RowNo + 1;
                                    ca.StartColumn = DuplicatesToWriteInColumnNo;
                                    ca.EndRow = wsIF.Cells.MaxRow;
                                    ca.EndColumn = DuplicatesToWriteInColumnNo;
                                    while (true)
                                    {
                                        cell = wsIF.Cells.FindStringContains(ExactDuplicateSetComment, cell, false, ca);
                                        if (cell == null)
                                            break;
                                        if (wsIF.Cells[cell.Row, DuplicatesToWriteInColumnNo].StringValue.Trim().ToUpper() == ExactDuplicateSetComment.ToUpper() &&
                                            wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].StringValue.Trim().ToUpper().Contains("CHECK"))
                                        {
                                            wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("Check");
                                            break;
                                        }
                                    }
                                }
                            }
                            #endregion

                            DuplicatesToWriteInColumnNo++;
                        }
                    }
                    #endregion

                    #region Normalized duplicates
                    if (model.IsToFindDuplicatesWithNormalizedMatch)
                    {
                        LastColumnNo = wsIF.Cells.MaxColumn;
                        DuplicatesToWriteInColumnNo = LastColumnNo + 1;
                        for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                        {
                            wsIF.Cells[0, DuplicatesToWriteInColumnNo].PutValue("Normalized duplicate on " + wsIF.Cells[0, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim());
                            for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                            {
                                CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                if (CellValueToFind.Length > 0)
                                {
                                    CellValueToFind = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFind);
                                    if (wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                                    {
                                        for (int r = RowNo + 1; r <= wsIF.Cells.MaxRow; r++)
                                        {
                                            CellValueToFindOn = wsIF.Cells[r, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                            CellValueToFindOn = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFindOn);
                                            if (CellValueToFind.ToUpper() == CellValueToFindOn.ToUpper())
                                                wsIF.Cells[r, DuplicatesToWriteInColumnNo].PutValue("Normalized Duplicate of " + model.UniqueIdentifierColumnName.Trim() + ":" + wsIF.Cells[RowNo, UIColumnNo].StringValue.Trim());
                                        }
                                    }
                                }
                            }
                            DuplicatesToWriteInColumnNo++;
                        }
                    }
                    #endregion

                    #region Selected Percentage of duplicates
                    if (model.IsToFindDuplicatesWithPercentageMatch)
                    {
                        LastColumnNo = wsIF.Cells.MaxColumn;
                        DuplicatesToWriteInColumnNo = LastColumnNo + 1;
                        CellValueToFind = ""; CellValueToFindOn = "";
                        for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                        {
                            wsIF.Cells[0, DuplicatesToWriteInColumnNo].PutValue( model.PercentageMatch.ToString() + "% duplicate on " + wsIF.Cells[0, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim());
                            for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                            {
                                CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                CellValueToFind = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFind);
                                if (CellValueToFind.Length > 4)
                                {
                                    CellValueToFind = GetAppliedPercentString(CellValueToFind, true, model.PercentageMatch);
                                    if (wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                                    {
                                        for (int r = RowNo + 1; r <= wsIF.Cells.MaxRow; r++)
                                        {
                                            CellValueToFindOn = wsIF.Cells[r, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                            CellValueToFindOn = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFindOn);
                                            if (CellValueToFindOn.Length > 4 && CellValueToFindOn.Length >= CellValueToFind.Length)
                                            {
                                                if (CellValueToFindOn.ToUpper().Contains(CellValueToFind.ToUpper()))
                                                    wsIF.Cells[r, DuplicatesToWriteInColumnNo].PutValue(model.PercentageMatch + "% duplicate of " + model.UniqueIdentifierColumnName + ":" + wsIF.Cells[RowNo, UIColumnNo].StringValue.Trim());
                                            }
                                        }
                                    }

                                    CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                    CellValueToFind = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFind);
                                    CellValueToFind = GetAppliedPercentString(CellValueToFind, false, model.PercentageMatch);
                                    if (wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                                    {
                                        for (int r = RowNo + 1; r <= wsIF.Cells.MaxRow; r++)
                                        {
                                            CellValueToFindOn = wsIF.Cells[r, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                            CellValueToFindOn = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFindOn);
                                            if (CellValueToFindOn.Length > 4 && CellValueToFindOn.Length >= CellValueToFind.Length)
                                            {
                                                if (CellValueToFindOn.ToUpper().Contains(CellValueToFind.ToUpper()))
                                                    wsIF.Cells[r, DuplicatesToWriteInColumnNo].PutValue(model.PercentageMatch.ToString() + "% duplicate of " + model.UniqueIdentifierColumnName.Trim() + ":" + wsIF.Cells[RowNo, UIColumnNo].StringValue.Trim());
                                            }
                                        }
                                    }
                                }
                            }
                            DuplicatesToWriteInColumnNo++;
                        }
                    }
                    #endregion
                }
                else
                {
                    if(model.ColumnsToCheckForDuplicates.Count == 1)
                    {
                        File.Delete(InputFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please select multiple columns to combine");
                    }

                    #region Exact duplicates
                    int combinedStringColumnNo = LastColumnNo + 1;
                    int ExactDuplicatesToWriteInColumnNo = combinedStringColumnNo + 1;
                    int CheckOrOKCommentColumnNo = ExactDuplicatesToWriteInColumnNo + 1;
                    string combinedString = string.Empty;
                    string DuplicateSetComment = string.Empty;
                    string UnMatchedColumnHeadingOrAttribute = string.Empty;
                    if (model.IsToFindDuplicatesWithExactMatch)
                    {
                        #region Writing combined string
                        wsIF.Cells[0, combinedStringColumnNo].PutValue("Combined String");
                        wsIF.Cells[0, ExactDuplicatesToWriteInColumnNo].PutValue("Exact Duplicate on combination of " + selectedColumnNames);
                        wsIF.Cells[0, CheckOrOKCommentColumnNo].PutValue("Check or OK Comment");
                        for (int RowNo = 1; RowNo <= wsIF.Cells.MaxRow; RowNo++)
                        {
                            combinedString = string.Empty;
                            for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                                combinedString += wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                            wsIF.Cells[RowNo, combinedStringColumnNo].PutValue(combinedString);
                        }
                        #endregion

                        #region Writing Duplicate Set
                        bool IsSetPrinted = false;
                        CellArea ca = new CellArea();
                        int DuplicateSetCounter = 1;
                        Aspose.Cells.Cell cell = wsIF.Cells[1, ExactDuplicatesToWriteInColumnNo];
                        for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                        {
                            if (string.IsNullOrEmpty(wsIF.Cells[RowNo, ExactDuplicatesToWriteInColumnNo].StringValue.Trim()))
                            {
                                combinedString = wsIF.Cells[RowNo, combinedStringColumnNo].StringValue.Trim();
                                IsSetPrinted = false;
                                ca.StartRow = RowNo + 1;
                                ca.StartColumn = combinedStringColumnNo;
                                ca.EndRow = wsIF.Cells.MaxRow;
                                ca.EndColumn = combinedStringColumnNo;
                                while (true)
                                {
                                    cell = wsIF.Cells.FindStringContains(combinedString, cell, false, ca);
                                    if (cell == null)
                                        break;

                                    if (wsIF.Cells[cell.Row, combinedStringColumnNo].StringValue.Trim().ToUpper() == combinedString.ToUpper())
                                    {
                                        wsIF.Cells[RowNo, ExactDuplicatesToWriteInColumnNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter.ToString());
                                        wsIF.Cells[cell.Row, ExactDuplicatesToWriteInColumnNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter.ToString());
                                        IsSetPrinted = true;
                                    }
                                }

                                if (IsSetPrinted)
                                    DuplicateSetCounter++;
                            }
                        }
                        #endregion

                        #region Writing Check or OK comment
                        cell = wsIF.Cells[1, ExactDuplicatesToWriteInColumnNo];
                        string ExactDuplicateSetComment = string.Empty;
                        bool AllColumnValuesAreSame = true;
                        for (int RowNo = 1; RowNo <= wsIF.Cells.MaxRow; RowNo++)
                        {
                            if (!string.IsNullOrEmpty(wsIF.Cells[RowNo, ExactDuplicatesToWriteInColumnNo].StringValue.Trim()))
                            {
                                ExactDuplicateSetComment = wsIF.Cells[RowNo, ExactDuplicatesToWriteInColumnNo].StringValue.Trim();
                                ca.StartRow = RowNo + 1;
                                ca.StartColumn = ExactDuplicatesToWriteInColumnNo;
                                ca.EndRow = wsIF.Cells.MaxRow;
                                ca.EndColumn = ExactDuplicatesToWriteInColumnNo;
                                while (true)
                                {
                                    cell = wsIF.Cells.FindStringContains(ExactDuplicateSetComment, cell, false, ca);
                                    if (cell == null)
                                        break;

                                    if (string.IsNullOrEmpty(wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].StringValue))
                                    {
                                        if (wsIF.Cells[cell.Row, ExactDuplicatesToWriteInColumnNo].StringValue.Trim().ToUpper() == ExactDuplicateSetComment.ToUpper())
                                        {
                                            AllColumnValuesAreSame = true;
                                            for (int col = 0; col <= LastColumnNo; col++)
                                            {
                                                if (wsIF.Cells[0, col].StringValue.Trim().ToLower() != "additional information" && col != UIColumnNo)
                                                {
                                                    if (wsIF.Cells[cell.Row, col].StringValue.Trim().ToLower() != wsIF.Cells[RowNo, col].StringValue.Trim().ToLower())
                                                    {
                                                        if (wsIF.Cells[0, col - 1].StringValue.Trim().ToLower().Contains("att") && wsIF.Cells[0, col].StringValue.Trim().ToLower().Contains("value"))
                                                            UnMatchedColumnHeadingOrAttribute = wsIF.Cells[cell.Row, col - 1].StringValue.Trim();
                                                        else
                                                            UnMatchedColumnHeadingOrAttribute = wsIF.Cells[0, col].StringValue.Trim();
                                                        
                                                        wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("Check");
                                                        wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].PutValue("Check - " + UnMatchedColumnHeadingOrAttribute);
                                                        AllColumnValuesAreSame = false;
                                                        break;
                                                    }
                                                }
                                            }

                                            if (AllColumnValuesAreSame)
                                            {
                                                wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("OK");
                                                wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].PutValue("OK");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Writing Check Comment to all set even if one row has Check Comment
                        for (int RowNo = 1; RowNo <= wsIF.Cells.MaxRow; RowNo++)
                        {
                            if (wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].StringValue.Trim().ToUpper() == "OK")
                            {
                                ExactDuplicateSetComment = wsIF.Cells[RowNo, ExactDuplicatesToWriteInColumnNo].StringValue.Trim();
                                ca.StartRow = RowNo + 1;
                                ca.StartColumn = ExactDuplicatesToWriteInColumnNo;
                                ca.EndRow = wsIF.Cells.MaxRow;
                                ca.EndColumn = ExactDuplicatesToWriteInColumnNo;
                                while (true)
                                {
                                    cell = wsIF.Cells.FindStringContains(ExactDuplicateSetComment, cell, false, ca);
                                    if (cell == null)
                                        break;
                                    if (wsIF.Cells[cell.Row, ExactDuplicatesToWriteInColumnNo].StringValue.Trim().ToUpper() == ExactDuplicateSetComment.ToUpper() &&
                                        wsIF.Cells[cell.Row, CheckOrOKCommentColumnNo].StringValue.Trim().ToUpper().Contains("CHECK"))
                                    {
                                        wsIF.Cells[RowNo, CheckOrOKCommentColumnNo].PutValue("Check");
                                        break;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion

                    #region Normalized duplicates
                    bool AllColumnValuesMatched = false;
                    bool AreAllColumnsEmpty = true;
                    string CellValueToFind = "", CellValueToFindOn = "";
                    if (model.IsToFindDuplicatesWithNormalizedMatch)
                    {
                        LastColumnNo = wsIF.Cells.MaxColumn;
                        DuplicatesToWriteInColumnNo = LastColumnNo + 1;
                        wsIF.Cells[0, DuplicatesToWriteInColumnNo].PutValue("Normalized duplicate on combination of " + selectedColumnNames);
                        for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                        {
                            if (wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                            {
                                for (int r = RowNo + 1; r <= wsIF.Cells.MaxRow; r++)
                                {
                                    AllColumnValuesMatched = false;
                                    AreAllColumnsEmpty = true;
                                    for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                                    {
                                        CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                        CellValueToFind = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFind);

                                        if (CellValueToFind.Length > 0)
                                        {
                                            CellValueToFindOn = wsIF.Cells[r, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                            CellValueToFindOn = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFindOn);

                                            if (CellValueToFindOn.Length > 0)
                                                AreAllColumnsEmpty = false;

                                            if (CellValueToFind.ToUpper() == CellValueToFindOn.ToUpper())
                                                AllColumnValuesMatched = true;
                                            else
                                            {
                                                AllColumnValuesMatched = false;
                                                break;
                                            }
                                        }
                                    }

                                    if (!AreAllColumnsEmpty)
                                    {
                                        if (AllColumnValuesMatched)
                                            wsIF.Cells[r, DuplicatesToWriteInColumnNo].PutValue("Normalized Duplicate of " + model.UniqueIdentifierColumnName.Trim() + ":" + wsIF.Cells[RowNo, UIColumnNo].StringValue.Trim());
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Selected Percentage of duplicates
                    if (model.IsToFindDuplicatesWithPercentageMatch)
                    {
                        LastColumnNo = wsIF.Cells.MaxColumn;
                        DuplicatesToWriteInColumnNo = LastColumnNo + 1;
                        CellValueToFind = ""; CellValueToFindOn = "";
                        wsIF.Cells[0, DuplicatesToWriteInColumnNo].PutValue(model.PercentageMatch.ToString() + "% duplicate on combination of " + selectedColumnNames);
                        for (int RowNo = 1; RowNo < wsIF.Cells.MaxRow; RowNo++)
                        {
                            if (wsIF.Cells[RowNo, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                            {
                                for (int r = RowNo + 1; r <= wsIF.Cells.MaxRow; r++)
                                {
                                    AllColumnValuesMatched = false;
                                    AreAllColumnsEmpty = true;
                                    for (int index = 0; index < arrSelectedColumnNos.Count(); index++)
                                    {
                                        CellValueToFind = wsIF.Cells[RowNo, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                        CellValueToFind = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFind);
                                        if (CellValueToFind.Length > 4)
                                        {
                                            CellValueToFind = GetAppliedPercentString(CellValueToFind, true,model.PercentageMatch);
                                            CellValueToFindOn = wsIF.Cells[r, Convert.ToInt32(arrSelectedColumnNos[index])].StringValue.Trim();
                                            CellValueToFindOn = GetStringValueWithoutSpecialCharactersAndLeadingZeros(CellValueToFindOn);

                                            if (CellValueToFindOn.Length > 4 && CellValueToFindOn.Length >= CellValueToFind.Length)
                                            {
                                                if (CellValueToFindOn.Length > 0)
                                                    AreAllColumnsEmpty = false;

                                                if (CellValueToFindOn.ToUpper().Contains(CellValueToFind.ToUpper()))
                                                    AllColumnValuesMatched = true;
                                                else
                                                {
                                                    AllColumnValuesMatched = false;
                                                    break;
                                                }

                                                if (!AreAllColumnsEmpty)
                                                {
                                                    if (AllColumnValuesMatched)
                                                    {
                                                        if (wsIF.Cells[r, DuplicatesToWriteInColumnNo].StringValue.Trim().Length == 0)
                                                            wsIF.Cells[r, DuplicatesToWriteInColumnNo].PutValue(model.PercentageMatch.ToString() + "% Duplicate of " + model.UniqueIdentifierColumnName.Trim() + ":" + wsIF.Cells[RowNo, UIColumnNo].StringValue.Trim());
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region Save and Download the file
                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + model.UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

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

        private string GetStringValueWithoutSpecialCharactersAndLeadingZeros(string InputValue)
        {
            string StringValueWithoutSpecialCharactersAndLeadingZeros = "";

            if (InputValue.Length > 0)
            {
                StringValueWithoutSpecialCharactersAndLeadingZeros = Regex.Replace(InputValue, @"[^0-9a-zA-Z]+", "");
                StringValueWithoutSpecialCharactersAndLeadingZeros = StringValueWithoutSpecialCharactersAndLeadingZeros.TrimStart('0');
            }

            return StringValueWithoutSpecialCharactersAndLeadingZeros.Trim();
        }

        private string GetAppliedPercentString(string InputString, bool IsFromLeft, double PercentageOfMatch)
        {
            int NoOfCharactersToRemove = 0;
            double f = 0.0, DecimalPart = 0.0;
            string AppliedPercentString = "";
            double RemainingPercentage = 0.0;

            RemainingPercentage = 100.0 - PercentageOfMatch;
            InputString = InputString.Trim();
            if (InputString.Length > 0)
            {
                f = (InputString.Length * RemainingPercentage) / 100.0;        //Remaining % characters to remove
                DecimalPart = f - Math.Truncate(f);
                if (DecimalPart <= 0.5)
                    NoOfCharactersToRemove = (int)Math.Truncate(f);
                else
                    NoOfCharactersToRemove = (int)Math.Ceiling(f);

                if (NoOfCharactersToRemove > 0)
                {
                    if (IsFromLeft)
                        AppliedPercentString = InputString.Substring(0, InputString.Length - NoOfCharactersToRemove);
                    else
                        AppliedPercentString = InputString.Substring(NoOfCharactersToRemove);
                }
                else
                    AppliedPercentString = InputString;
            }

            return AppliedPercentString;
        }
    }
}
