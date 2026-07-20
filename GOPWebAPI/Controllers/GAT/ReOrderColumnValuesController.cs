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
    [RoutePrefix("api/ReOrderColumnValues")]
    public class ReOrderColumnValuesController : ApiController
    {
        #region Validate the input data
        [HttpPost]
        [Route("ValidateInputData")]
        public HttpResponseMessage ValidateInputData(string FileName, string ColumnToSort, string Delimiter)
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
                Worksheet wsIF = wbIF.Worksheets[0];

                #region checking unique column headings
                string ColumnName = "", DuplicateColumnName = "";
                int ColumnHeadingCounter = 0;
                for (int ColNo = 0; ColNo <= wsIF.Cells.MaxColumn; ColNo++)
                {
                    ColumnName = wsIF.Cells[0, ColNo].StringValue.Trim();
                    for (int c = 0; c <= wsIF.Cells.MaxColumn; c++)
                    {
                        if (ColumnName.ToUpper() == wsIF.Cells[0, c].StringValue.Trim().ToUpper())
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
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column headings from input first worksheet must be unique. Duplicate column name:'" + DuplicateColumnName + "'");
                    }
                    ColumnHeadingCounter = 0;
                }
                #endregion

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }

                if(string.IsNullOrEmpty(Delimiter))
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Delimiter");
                }

                if (string.IsNullOrEmpty(ColumnToSort))
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Column to sort");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Input data validated successfully");
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

        #region Sort selected Column Values from input file
        [HttpPost]
        [Route("SortTheValuesFromSelectedColumn")]
        public HttpResponseMessage SortTheValuesFromSelectedColumn(string UploadedInputFileName, string InputFileName, string ColumnToSort, string Delimiter, bool IsToWriteTYPEAttributeAtTheBeginning = true)
        {
            int LastColumnNo, CommentColumnNo;
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Finding the column no. to sort
                int SelectedColumnNo = 0;
                for (int col = 0; col <= wsIF.Cells.MaxColumn; col++)
                {
                    if (wsIF.Cells[0, col].StringValue.Trim().ToLower() == ColumnToSort.Trim().ToLower())
                    {
                        SelectedColumnNo = col;
                        break;
                    }
                }
                #endregion

                #region Writing new column heading
                LastColumnNo = wsIF.Cells.MaxColumn + 1;
                wsIF.Cells[0, LastColumnNo].PutValue("Re-Ordered " + ColumnToSort);
                CommentColumnNo = LastColumnNo + 1;
                wsIF.Cells[0, CommentColumnNo].PutValue("Comment");
                #endregion

                #region Sorting and Writing the Data
                string stringToSort = "";
                string[] ArrayOfSplittedValues;
                string[] ArrayOfSplittedValuesBasedOnColon;
                string[] ArrayOfSplittedValuesBasedDelimiter;
                string SortedString = "";
                for (int RowNo = 1; RowNo <= wsIF.Cells.MaxRow; RowNo++)
                {
                    stringToSort = wsIF.Cells[RowNo, SelectedColumnNo].StringValue.Trim();

                    if (stringToSort.Split(':').Count() >= 2)
                    {
                        ArrayOfSplittedValuesBasedOnColon = stringToSort.Split(':');
                        foreach (string value in ArrayOfSplittedValuesBasedOnColon)
                        {
                            ArrayOfSplittedValuesBasedDelimiter = value.Split(new[] { Delimiter }, StringSplitOptions.None);
                            if (ArrayOfSplittedValuesBasedDelimiter.Count() > 2)
                            {
                                wsIF.Cells[RowNo, CommentColumnNo].PutValue("Multiple values to attribute");
                                break;
                            }
                        }
                    }

                    ArrayOfSplittedValues = stringToSort.Split(new[] { Delimiter }, StringSplitOptions.None);

                    Array.Sort(ArrayOfSplittedValues);
                    SortedString = "";

                    if (IsToWriteTYPEAttributeAtTheBeginning)
                    {
                        foreach (string value in ArrayOfSplittedValues)
                        {
                            if (value.ToUpper().Contains("TYPE:") || value.ToUpper().Contains("TYPE: "))
                            {
                                if (SortedString.Length > 0)
                                    SortedString = SortedString + Delimiter + value;
                                else
                                    SortedString = value;
                            }
                        }

                        foreach (string value in ArrayOfSplittedValues)
                        {
                            if (value.Contains(':'))
                            {
                                if (!(value.ToUpper().Contains("TYPE:") && value.ToUpper().Contains("TYPE: ")))
                                {
                                    if (SortedString.Length > 0)
                                        SortedString = SortedString + Delimiter + value;
                                    else
                                        SortedString = value;
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (string value in ArrayOfSplittedValues)
                        {
                            if (value.Contains(':'))
                            {
                                if (SortedString.Length > 0)
                                    SortedString = SortedString + Delimiter + value;
                                else
                                    SortedString = value;
                            }
                        }
                    }

                    foreach (string value in ArrayOfSplittedValues)
                    {
                        if (!value.Contains(':'))
                        {
                            if (value.StartsWith("0") || value.StartsWith("1") || value.StartsWith("2") || value.StartsWith("3") || value.StartsWith("4") ||
                                value.StartsWith("5") || value.StartsWith("6") || value.StartsWith("7") || value.StartsWith("8") || value.StartsWith("9"))
                            {
                                if (SortedString.Length > 0)
                                    SortedString = SortedString + Delimiter + value;
                                else
                                    SortedString = value;
                            }
                        }
                    }

                    foreach (string value in ArrayOfSplittedValues)
                    {
                        if (!value.Contains(':'))
                        {
                            if (!value.StartsWith("0") && !value.StartsWith("1") && !value.StartsWith("2") && !value.StartsWith("3") && !value.StartsWith("4") &&
                                !value.StartsWith("5") && !value.StartsWith("6") && !value.StartsWith("7") && !value.StartsWith("8") && !value.StartsWith("9"))
                            {
                                if (SortedString.Length > 0)
                                    SortedString = SortedString + Delimiter + value;
                                else
                                    SortedString = value;
                            }
                        }
                    }

                    wsIF.Cells[RowNo, LastColumnNo].PutValue(SortedString);
                }
                #endregion

                #region Save the file
                wsIF.AutoFitColumns();
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
