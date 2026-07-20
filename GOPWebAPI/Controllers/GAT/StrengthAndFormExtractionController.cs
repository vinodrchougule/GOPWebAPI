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
    [RoutePrefix("api/StrengthAndFormExtraction")]
    public class StrengthAndFormExtractionController : ApiController
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

                #region Checking worksheet count in workbook
                if (wbIF.Worksheets.Count < 4)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file should have minimum four worksheets.");
                }
                #endregion

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "EXAMPLE VALUES (INPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Example Values (Input)' not found in input file first worksheet. Please select a valid file.");
                }
                #endregion

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                if (ws2.Cells[0, 0].StringValue.Trim().ToUpper() != "UOM/UNITS/STRENGTHS(MOST COMMONLY OCCURING) ( INPUT & OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'UOM/Units/Strengths(Most commonly occuring) ( Input & Output)' not found in input file second worksheet. Please select a valid file.");
                }
                #endregion

                #region Validating the third worksheet
                Worksheet ws3 = wbIF.Worksheets[2];
                if (ws3.Cells[0, 0].StringValue.Trim().ToUpper() != "MOST COMMON FORMS (INPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Most Common Forms (Input)' not found in input file second worksheet. Please select a valid file.");
                }
                #endregion

                #region Validating the fourth worksheet
                Worksheet ws4 = wbIF.Worksheets[3];
                if (ws4.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO.")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Sl No.' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 1].StringValue.Trim().ToUpper() != "INPUT DESC (INPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Input Desc (Input)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 2].StringValue.Trim().ToUpper() != "STRENGTH1 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Strength1 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 3].StringValue.Trim().ToUpper() != "STRENGTH2 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Strength2 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 4].StringValue.Trim().ToUpper() != "STRENGTH3 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Strength3 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 5].StringValue.Trim().ToUpper() != "FORM1 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'Form1 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 6].StringValue.Trim().ToUpper() != "FORM2 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [G1] with Value 'Form2 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 7].StringValue.Trim().ToUpper() != "FORM3 (OUTPUT)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [H1] with Value 'Form3 (Output)' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File fourth Worksheet has no data rows.");
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
        [Route("ExtractStrengthsAndFormsAndWriteToOutput")]
        public HttpResponseMessage ExtractStrengthsAndFormsAndWriteToOutput(string UploadedInputFileName, string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                string ExampleValueString = string.Empty, MostCommonUOM = string.Empty, InputDescription = string.Empty, Strength = string.Empty;
                int LastEmptyStrengthColNo = 2;

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet wsEV = wbIF.Worksheets[0];          //Example Values
                int maxEVRows = wsEV.Cells.MaxRow;

                Worksheet wsMCV = wbIF.Worksheets[1];          //Most Common Strengths or UOMs
                int maxMCVRows = wsMCV.Cells.MaxRow;

                Worksheet wsMCF = wbIF.Worksheets[2];          //Most Common Forms
                int maxMCFRows = wsMCF.Cells.MaxRow;

                Worksheet wsID = wbIF.Worksheets[3];          //Input Description
                int maxIDRows = wsID.Cells.MaxRow;

                bool IsEVUOMExists = false;
                for (int evr = 1; evr <= maxEVRows; evr++)
                {
                    #region Searching Example Value UOM in Most Common UOMs and appending
                    ExampleValueString = wsEV.Cells[evr, 0].StringValue.Trim().Replace('/', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace('+', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace('-', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace('(', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace(')', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace(',', ' ');
                    ExampleValueString = ExampleValueString.Trim().Replace('&', ' ');
                    string[] arrExampleValue = ExampleValueString.Split(' ');
                    foreach (string ev in arrExampleValue)
                    {
                        string replacedEV = Regex.Replace(ev, "[0-9.]", "").Trim();
                        if (replacedEV.ToUpper() != "MIU" &&
                            replacedEV.ToUpper() != "MIC" &&
                            replacedEV.ToUpper() != "LAKH" &&
                            replacedEV.ToUpper() != "MM" &&
                            replacedEV.ToUpper() != "LIQ" &&
                            replacedEV.ToUpper() != "U" &&
                            replacedEV.ToUpper() != "INJ")
                        {
                            if (replacedEV.Length > 0)
                            {
                                IsEVUOMExists = false;
                                if (maxMCVRows > 0)
                                {
                                    for (int mcr = 1; mcr <= maxMCVRows; mcr++)
                                    {
                                        if (wsMCV.Cells[mcr, 0].StringValue.Trim().ToUpper() == replacedEV.Trim().ToUpper())
                                        {
                                            IsEVUOMExists = true;
                                            break;
                                        }
                                    }

                                    if (!IsEVUOMExists)
                                    {
                                        maxMCVRows++;
                                        wsMCV.Cells[maxMCVRows, 0].PutValue(replacedEV);
                                        break;
                                    }
                                }
                                else
                                {
                                    maxMCVRows++;
                                    wsMCV.Cells[maxMCVRows, 0].PutValue(replacedEV);
                                    break;
                                }
                            }
                        }
                    }
                    #endregion
                }

                #region Searching most common value in Input Description and Writing Strength1, Strength2, Strength3
                char LeftSideChar, RightSideChar = 'a';
                string ValueBeforeUOM = string.Empty;
                for (int mcr = 1; mcr <= maxMCVRows; mcr++)
                {
                    MostCommonUOM = wsMCV.Cells[mcr, 0].StringValue.Trim();
                    for (int idr = 1; idr <= maxIDRows; idr++)
                    {
                        LastEmptyStrengthColNo = 2;
                        InputDescription = wsID.Cells[idr, 1].StringValue.Trim();
                        //loop as long as most common UOM is in Input String and not at the beginning
                        while (InputDescription.ToUpper().IndexOf(MostCommonUOM.ToUpper()) > 0)
                        {
                            int idxMCUOM = InputDescription.ToUpper().IndexOf(MostCommonUOM);

                            //Checking left side character should not be a letter
                            LeftSideChar = InputDescription[idxMCUOM - 1];

                            //Left side character must not be letter
                            if (!Char.IsLetter(LeftSideChar) && LeftSideChar.ToString() != "'")
                            {
                                //Checking right side character should not be a letter
                                if (idxMCUOM + MostCommonUOM.Length < InputDescription.Length)
                                    RightSideChar = InputDescription[idxMCUOM + MostCommonUOM.Length];

                                //if right side is not a letter or it is at the end of input string
                                if (!Char.IsLetter(RightSideChar) || (idxMCUOM + MostCommonUOM.Length == InputDescription.Length))
                                {
                                    #region Extract the strength and write it
                                    //Search from UOM found position till beginning
                                    for (int i = idxMCUOM; i >= 0; i--)
                                    {
                                        // if left side first character is space then continue
                                        if (InputDescription.Substring(idxMCUOM - 1, 1) == " " && i == idxMCUOM - 1)
                                            continue;

                                        //if index reaches to beginning or found the space character at the search index
                                        if (i == 0 || InputDescription.Substring(i, 1) == " ")
                                        {
                                            //get the strength
                                            if (i == 0)
                                            {
                                                Strength = InputDescription.Substring(i, (idxMCUOM - i) + MostCommonUOM.Length);
                                            }
                                            else
                                            {
                                                Strength = InputDescription.Substring(i + 1, (idxMCUOM - (i + 1)) + MostCommonUOM.Length);
                                            }

                                            //find the empty strength column no.
                                            while (!string.IsNullOrEmpty(wsID.Cells[idr, LastEmptyStrengthColNo].StringValue.Trim()))
                                                LastEmptyStrengthColNo++;

                                            //write the strength by replacing space with null
                                            Strength = Strength.Replace("-", "");
                                            Strength = Strength.Replace(" ", "");
                                            Strength = Strength.Replace("(", "");
                                            Strength = Strength.Replace(")", "");
                                            if (LastEmptyStrengthColNo <= 4)
                                                wsID.Cells[idr, LastEmptyStrengthColNo].PutValue(Strength);
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                            }

                            //Set Input Description to remaining string after removing the processed part of input description
                            if (InputDescription.Length - (idxMCUOM + MostCommonUOM.Length) > 0)
                                InputDescription = InputDescription.Substring(idxMCUOM + MostCommonUOM.Length, InputDescription.Length - (idxMCUOM + MostCommonUOM.Length));
                            else //if it was at the end then break the loop
                                break;
                        }
                    }
                }
                #endregion

                #region if strength1 contains % and strength2 contains MG or ML then replace strength1 with strength2 and replace strength2 with strength3
                for (int idr = 1; idr <= maxIDRows; idr++)
                {
                    if (wsID.Cells[idr, 2].StringValue.Trim().Contains("%") &&
                        (wsID.Cells[idr, 3].StringValue.Trim().ToUpper().Contains("MG") ||
                         wsID.Cells[idr, 3].StringValue.Trim().ToUpper().Contains("GM") ||
                         wsID.Cells[idr, 3].StringValue.Trim().ToUpper().Contains("ML")))
                    {
                        //Replace strength1 with strength2
                        wsID.Cells[idr, 2].PutValue(wsID.Cells[idr, 3].StringValue);
                        //Replace strength2 with strength3
                        wsID.Cells[idr, 3].PutValue(wsID.Cells[idr, 4].StringValue);
                    }

                    if (wsID.Cells[idr, 2].StringValue.Trim().Contains("%") &&
                        (wsID.Cells[idr, 4].StringValue.Trim().ToUpper().Contains("ML") ||
                         wsID.Cells[idr, 4].StringValue.Trim().ToUpper().Contains("GM")))
                    {
                        //Replace strength1 with strength3
                        wsID.Cells[idr, 2].PutValue(wsID.Cells[idr, 4].StringValue);
                        wsID.Cells[idr, 4].PutValue("");
                    }
                }
                #endregion

                #region Replace alphabets with blank from left to right until digit occurs
                string Strength1 = string.Empty, Strength2 = string.Empty, Strength3 = string.Empty;
                for (int idr = 1; idr <= maxIDRows; idr++)
                {

                    //Strength1
                    if (wsID.Cells[idr, 2].StringValue.Trim().Length > 0)
                    {
                        Strength1 = wsID.Cells[idr, 2].StringValue.Trim();
                        while (Strength1.Length > 0)
                        {
                            if (!Char.IsDigit(Strength1[0]))
                                Strength1 = ReplaceAt(Strength1, 0, 1, "");
                            else
                                break;
                        }
                        wsID.Cells[idr, 2].PutValue(Strength1);
                    }

                    //Strength2
                    if (wsID.Cells[idr, 3].StringValue.Trim().Length > 0)
                    {
                        Strength2 = wsID.Cells[idr, 3].StringValue.Trim();
                        while (Strength2.Length > 0)
                        {
                            if (!Char.IsDigit(Strength2[0]))
                                Strength2 = ReplaceAt(Strength2, 0, 1, "");
                            else
                                break;
                        }
                        wsID.Cells[idr, 3].PutValue(Strength2);
                    }

                    //Strength3
                    if (wsID.Cells[idr, 4].StringValue.Trim().Length > 0)
                    {
                        Strength3 = wsID.Cells[idr, 4].StringValue.Trim();
                        while (Strength3.Length > 0)
                        {
                            if (!Char.IsDigit(Strength3[0]))
                                Strength3 = ReplaceAt(Strength3, 0, 1, "");
                            else
                                break;
                        }
                        wsID.Cells[idr, 4].PutValue(Strength3);
                    }
                }
                #endregion

                #region if strength1 is blank, replace it with strength2 and if strength2 is blank, replace it with strength3
                for (int idr = 1; idr <= maxIDRows; idr++)
                {
                    if (string.IsNullOrEmpty(wsID.Cells[idr, 2].StringValue.Trim()) &&    //strength1
                        !string.IsNullOrEmpty(wsID.Cells[idr, 3].StringValue.Trim()))   //strength2
                    {
                        wsID.Cells[idr, 2].PutValue(wsID.Cells[idr, 3].StringValue.Trim());
                        wsID.Cells[idr, 3].PutValue("");
                    }

                    if (string.IsNullOrEmpty(wsID.Cells[idr, 3].StringValue.Trim()) &&    //strength2
                        !string.IsNullOrEmpty(wsID.Cells[idr, 4].StringValue.Trim()))   //strength3
                    {
                        wsID.Cells[idr, 3].PutValue(wsID.Cells[idr, 4].StringValue.Trim());
                        wsID.Cells[idr, 4].PutValue("");
                    }
                }
                #endregion

                #region Form Extraction
                string MostCommonForm = string.Empty;
                int LastEmptyFormColNo = 5;
                for (int mcf = 1; mcf <= maxMCFRows; mcf++)
                {
                    MostCommonForm = wsMCF.Cells[mcf, 0].StringValue.Trim();
                    for (int idr = 1; idr <= maxIDRows; idr++)
                    {
                        LastEmptyFormColNo = 5;
                        InputDescription = wsID.Cells[idr, 1].StringValue.Trim();
                        //if most common Form is in Input String
                        if (InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) >= 0)
                        {
                            //at the beginning
                            if (InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) == 0)
                            {
                                if (InputDescription.Length == MostCommonForm.Length ||
                                    (!Char.IsLetter(InputDescription[MostCommonForm.Length])))
                                {
                                    wsID.Cells[idr, LastEmptyFormColNo].PutValue(MostCommonForm);
                                    LastEmptyFormColNo++;
                                }
                            }

                            //in middle
                            if (InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) > 0)
                            {
                                if (InputDescription.Length > (InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) + MostCommonForm.Length))
                                {
                                    if (!Char.IsLetter(InputDescription[InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) - 1]) &&
                                        !Char.IsLetter(InputDescription[InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) + MostCommonForm.Length]))
                                    {
                                        if (MostCommonForm.ToUpper() != wsID.Cells[idr, 5].StringValue.Trim().ToUpper())
                                        {
                                            wsID.Cells[idr, LastEmptyFormColNo].PutValue(MostCommonForm);
                                            LastEmptyFormColNo++;
                                        }
                                    }
                                }
                            }

                            //at the end
                            if (InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) > 0 &&
                               InputDescription.ToUpper().EndsWith(MostCommonForm.ToUpper()))
                            {
                                if (!Char.IsLetter(InputDescription[InputDescription.ToUpper().IndexOf(MostCommonForm.ToUpper()) - 1]))
                                {
                                    if (MostCommonForm.ToUpper() != wsID.Cells[idr, 5].StringValue.Trim().ToUpper() &&
                                        MostCommonForm.ToUpper() != wsID.Cells[idr, 6].StringValue.Trim().ToUpper())
                                    {
                                        wsID.Cells[idr, LastEmptyFormColNo].PutValue(MostCommonForm);
                                        LastEmptyFormColNo++;
                                    }
                                }
                            }
                        }
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

        #region Helping Functions
        private string ReplaceAt(string str, int index, int length, string replace)
        {
            return str.Remove(index, Math.Min(length, str.Length - index))
                    .Insert(index, replace);
        }
        #endregion
    }
}
