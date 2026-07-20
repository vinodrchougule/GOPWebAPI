using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/FindMissingWords")]
    public class FindMissingWordsController : ApiController
    {
        #region Read Column Names From File
        [HttpGet]
        [Route("ReadColumnNamesFromFile")]
        public IHttpActionResult ReadColumnNamesFromFile(string FileName)
        {
            try
            {
                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
                List<string> ColumnNamesList = new List<string>();

                Workbook IFwb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilepath);
                Worksheet IFws = IFwb.Worksheets[0];

                for (int col = 0; col <= IFws.Cells.MaxColumn; col++)
                    ColumnNamesList.Add(IFws.Cells[0, col].StringValue.Trim());

                return Ok(ColumnNamesList);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate the input file data
        [HttpPost]
        [Route("ValidateInputFileData")]
        public HttpResponseMessage ValidateInputFileData([FromBody] FindMissingWords findMissingWords)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + findMissingWords.FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }

                #region checking unique column headings
                string ColumnName = string.Empty, DuplicateColumnName = string.Empty;
                for (int ColNo = 0; ColNo <= wsIF.Cells.MaxColumn; ColNo++)
                {
                    ColumnName = wsIF.Cells[0, ColNo].StringValue.Trim();

                    if(string.IsNullOrEmpty(ColumnName))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column headings from input first worksheet cannot be empty / blank");
                    }

                    for (int c = 0; c <= wsIF.Cells.MaxColumn; c++)
                    {
                        if (c == ColNo)
                            continue;

                        if (ColumnName.ToUpper() == wsIF.Cells[0, c].StringValue.Trim().ToUpper())
                        {
                            DuplicateColumnName = ColumnName;
                            break;
                        }
                    }

                    if(!string.IsNullOrEmpty(DuplicateColumnName))
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column headings from input first worksheet must be unique. Duplicate column name:'" + DuplicateColumnName + "'");
                    }
                }
                #endregion

                if(findMissingWords.InputColumns.Count <= 0) 
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please select the input column(s)");
                }

                #region Check each column from input columns exists in file
                bool IsColumnExists = false;
                foreach (string ic in findMissingWords.InputColumns) 
                {
                    IsColumnExists = false;
                    for (int c = 0; c <= wsIF.Cells.MaxColumn; c++)
                    {
                        if(wsIF.Cells[0, c].StringValue.ToUpper() == ic.ToUpper())
                        {
                            IsColumnExists = true;
                            break;
                        }
                    }

                    if(!IsColumnExists)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input column '" + ic + "' not found in uploaded file.");
                    }
                }
                #endregion

                #region Check each column from output columns exists in file
                IsColumnExists = false;
                foreach (string oc in findMissingWords.OutputColumns)
                {
                    IsColumnExists = false;
                    for (int c = 0; c <= wsIF.Cells.MaxColumn; c++)
                    {
                        if (wsIF.Cells[0, c].StringValue.ToUpper() == oc.ToUpper())
                        {
                            IsColumnExists = true;
                            break;
                        }
                    }

                    if (!IsColumnExists)
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Output column '" + oc + "' not found in uploaded file.");
                    }
                }
                #endregion

                if (findMissingWords.OutputColumns.Count <= 0)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please select the output column(s)");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Input file data validated Successfully.");
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

        #region Find Missing Words, Repeated Words and, New Words and Write To Excel output file
        [HttpPost]
        [Route("FindMissingRepeatedAndNewWordsAndWriteToExcel")]
        public HttpResponseMessage FindMissingRepeatedAndNewWordsAndWriteToExcel([FromBody] FindMissingWords findMissingWords)
        {
            int InputColumnNo, OutputColumnNo, MissingWordsColumnNo, RepeatedWordsColumnNo, NewWordsColumnNo;
            string inputString;
            string InputColumnsStringData, OutputColumnsStringData;
            bool IsWordExistInOutputColumn;
            string MissingWords, RepeatedWords, NewWords;
            List<int> InputColNos = new List<int>(); List<int> OutputColNos = new List<int>();

            try
            {
                string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + findMissingWords.FileName);

                Workbook IFwb = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                IFwb.Open(InputFilepath);
                Worksheet IFws = IFwb.Worksheets[0];
                int maxRow = IFws.Cells.MaxRow;

                #region Writing Headers
                InputColumnNo = IFws.Cells.MaxColumn + 1;
                IFws.Cells[0, InputColumnNo].PutValue("Input");
                
                OutputColumnNo = InputColumnNo + 1;
                IFws.Cells[0, OutputColumnNo].PutValue("Output");
                
                MissingWordsColumnNo = OutputColumnNo + 1;
                IFws.Cells[0, MissingWordsColumnNo].PutValue("Missing Words");

                RepeatedWordsColumnNo = MissingWordsColumnNo + 1;
                IFws.Cells[0, RepeatedWordsColumnNo].PutValue("Repeated Words (in output columns)");

                NewWordsColumnNo = RepeatedWordsColumnNo + 1;
                IFws.Cells[0, NewWordsColumnNo].PutValue("New Words (in output columns)");
                #endregion

                #region Add Input Column Nos to List
                foreach (string iColName in findMissingWords.InputColumns)
                {
                    for (int ColNo = 0; ColNo <= IFws.Cells.MaxColumn; ColNo++)
                    {
                        if (iColName.Trim().ToLower() == IFws.Cells[0, ColNo].StringValue.Trim().ToLower())
                        {
                            InputColNos.Add(ColNo);
                            break;
                        }
                    }
                }
                #endregion

                #region Add Output Column Nos to List
                foreach (string oColName in findMissingWords.OutputColumns)
                {
                    for (int ColNo = 0; ColNo <= IFws.Cells.MaxColumn; ColNo++)
                    {
                        if (oColName.Trim().ToLower() == IFws.Cells[0, ColNo].StringValue.Trim().ToLower())
                        {
                            OutputColNos.Add(ColNo);
                            break;
                        }
                    }
                }
                #endregion

                #region Starting to Process each row from file
                for (int row = 1; row <= maxRow; row++)
                {
                    MissingWords = string.Empty;
                    inputString = string.Empty;  InputColumnsStringData = string.Empty; OutputColumnsStringData = string.Empty;

                    #region Concatenate input columns data
                    foreach (int iColNo in InputColNos)
                    {
                        inputString = inputString + IFws.Cells[row, iColNo].StringValue + "ʬ";
                        InputColumnsStringData = InputColumnsStringData + IFws.Cells[row, iColNo].StringValue + " ";
                    }
                    #endregion

                    #region Concatenate output columns data
                    foreach (int oColNo in OutputColNos)
                        OutputColumnsStringData = OutputColumnsStringData + IFws.Cells[row, oColNo].StringValue + " ";
                    #endregion

                    #region Replace Special Characters with Unicode Character
                    inputString = inputString.Replace('X', 'ʬ');
                    inputString = inputString.Replace('x', 'ʬ');
                    inputString = inputString.Replace(' ', 'ʬ');
                    inputString = inputString.Replace(',', 'ʬ');
                    inputString = inputString.Replace('&', 'ʬ');
                    inputString = inputString.Replace(':', 'ʬ');
                    inputString = inputString.Replace('(', 'ʬ');
                    inputString = inputString.Replace(')', 'ʬ');
                    inputString = inputString.Replace('{', 'ʬ');
                    inputString = inputString.Replace('}', 'ʬ');
                    inputString = inputString.Replace('[', 'ʬ');
                    inputString = inputString.Replace(']', 'ʬ');
                    inputString = inputString.Replace(';', 'ʬ');
                    inputString = inputString.Replace('|', 'ʬ');
                    inputString = inputString.Replace('"', 'ʬ');
                    #endregion

                    #region Get input distinct words
                    string[] inputWordsArray = inputString.Split('ʬ');

                    string[] inputDistinctWordsArray = inputWordsArray.Distinct().ToArray();
                    #endregion

                    #region Finding missing words
                    foreach (string word in inputDistinctWordsArray)
                    {
                        string NewWord = word;
                        if (NewWord.Contains('.') && (!NewWord.ToUpper().Contains('X')))
                        {
                            char left1char;
                            if (NewWord.IndexOf('.') > 1)
                            {
                                left1char = Convert.ToChar(NewWord.Substring(NewWord.IndexOf('.') - 1, 1));
                                int _asciiOfleft1char = char.ConvertToUtf32(left1char.ToString(), 0);
                                if (_asciiOfleft1char < 48 || _asciiOfleft1char > 57)
                                    NewWord = Regex.Replace(NewWord, "[.]", "");
                            }
                        }

                        IsWordExistInOutputColumn = false;
                        foreach (var oColNo in OutputColNos)
                        {
                            if (IFws.Cells[row, oColNo].StringValue.ToUpper().Contains(NewWord.Trim().ToUpper()))
                            {
                                IsWordExistInOutputColumn = true;
                                break;
                            }
                        }

                        if (!IsWordExistInOutputColumn)
                        {
                            if (MissingWords.Length > 0)
                            {
                                if (!MissingWords.ToUpper().Contains(NewWord.ToUpper()))
                                    MissingWords = MissingWords + "," + NewWord;
                            }
                            else
                                MissingWords = NewWord;
                        }
                    }
                    #endregion

                    #region Finding Repeated Words
                    int WordRepeatedCount = 0;
                    OutputColumnsStringData = OutputColumnsStringData.Replace(": ", " ");
                    OutputColumnsStringData = OutputColumnsStringData.Replace(", ", " ");
                    OutputColumnsStringData = OutputColumnsStringData.Replace("; ", " ");
                    OutputColumnsStringData = OutputColumnsStringData.Replace(";", " ");
                    OutputColumnsStringData = OutputColumnsStringData.Replace("/", " ");
                    //OutputColumnsStringData = OutputColumnsStringData.Replace("-", " ");
                    OutputColumnsStringData = OutputColumnsStringData.Replace("|", " ");

                    string[] outputWordsArray = OutputColumnsStringData.Split(' ');
                    string[] outputDistinctWordsArray = outputWordsArray.Distinct().ToArray();

                    RepeatedWords = "";
                    foreach (string dWord in outputDistinctWordsArray)
                    {
                        if (dWord.Trim().Length > 0)
                        {
                            WordRepeatedCount = 0;
                            foreach (string oWord in outputWordsArray)
                            {
                                if (dWord.ToUpper() == oWord.ToUpper())
                                {
                                    WordRepeatedCount++;
                                    if (WordRepeatedCount > 1)
                                    {
                                        if (RepeatedWords.Length > 0)
                                            RepeatedWords = RepeatedWords + "," + dWord.Trim();
                                        else
                                            RepeatedWords = dWord.Trim();

                                        break;
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Finding New Words
                    NewWords = "";
                    foreach (string oWord in outputDistinctWordsArray)
                    {
                        if (oWord.Trim().Length > 0)
                        {
                            if (!InputColumnsStringData.ToUpper().Contains(oWord.ToUpper()))
                            {
                                if (NewWords.Trim().Length > 0)
                                    NewWords = NewWords + "," + oWord.Trim();
                                else
                                    NewWords = oWord.Trim();
                            }
                        }
                    }
                    #endregion

                    #region Write the output
                    IFws.Cells[row, InputColumnNo].PutValue(InputColumnsStringData);
                    IFws.Cells[row, OutputColumnNo].PutValue(OutputColumnsStringData);
                    if (MissingWords.Trim().Length > 0)
                        IFws.Cells[row, MissingWordsColumnNo].PutValue(MissingWords);
                    if (RepeatedWords.Trim().Length > 0)
                        IFws.Cells[row, RepeatedWordsColumnNo].PutValue(RepeatedWords);
                    if (NewWords.Trim().Length > 0)
                        IFws.Cells[row, NewWordsColumnNo].PutValue(NewWords);
                    #endregion
                }
                #endregion

                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                //IFws.AutoFitColumns();        //throwing "A generic error occurred in GDI+." need to install few fonts on server / set fonts folder
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + findMissingWords.FileName;
                IFwb.Save(filename);

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

        #region Download the output file
        [HttpGet]
        [Route("DownloadOutputFile")]
        public HttpResponseMessage DownloadOutputFile(string FileFullName)
        {
            try
            {
                string FileName = Path.GetFileName(FileFullName.Replace("\\\\", "\\"));
                byte[] bytes = File.ReadAllBytes(FileFullName);

                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                response.Content = new ByteArrayContent(bytes);

                response.Content.Headers.ContentLength = bytes.LongLength;

                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileFullName));

                return response;
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
