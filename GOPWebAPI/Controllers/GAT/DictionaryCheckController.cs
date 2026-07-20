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
using System.Threading;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/DictionaryCheck")]
    public class DictionaryCheckController : ApiController
    {
        #region Validate the input uploaded file
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
                Worksheet wsIFFW = wbIF.Worksheets[0];      //Input File First Worksheet

                #region Check uploaded file first worksheet has columns as per template
                if (wsIFFW.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 1].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Noun'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 2].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCol = 1; aCol <= 99; aCol += 2)
                {
                    hdr = wsIFFW.Cells[0, aCol + 2].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCol + 3).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' .Please download format from Template.");
                    }

                    hdr = wsIFFW.Cells[0, aCol + 3].StringValue.Trim().ToUpper();
                    if (hdr != "VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCol + 4).ToString() + "] must contain 'Value" + attNo.ToString() + "' .Please download format from Template.");
                    }
                    attNo++;
                }

                if (wsIFFW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "success");
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

        #region Validate the dictionary uploaded file
        [HttpPost]
        [Route("ValidateDictionaryFile")]
        public HttpResponseMessage ValidateDictionaryFile(string FileName)
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid dictionary file of xlsx format only");

                Workbook wbDF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbDF.Open(UploadedFilepath);
                Worksheet wsDFFW = wbDF.Worksheets[0];      //Dictionary File First Worksheet

                #region Check uploaded file first worksheet has columns as per template
                if (wsDFFW.Cells[0, 0].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Noun' not found in dictionary file first worksheet. Please select a valid file.");
                }

                if (wsDFFW.Cells[0, 1].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Modifier' not found in dictionary file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1, aCol = 1;
                while(attNo<=50)
                {
                    hdr = wsDFFW.Cells[0, aCol + 1].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCol + 2).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' .Please download format from Template.");
                    }

                    attNo++; aCol++;
                }

                if (wsDFFW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Dictionary File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "success");
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

        #region Find the Noun,Modifier, Attribute from the input file not found in dictionary and write to second worksheet
        [HttpPost]
        [Route("WriteUnMatchedNMAs")]
        public HttpResponseMessage WriteUnMatchedNMAs(string UploadedInputFileName,string InputFileName, string DictionaryFileName)
        {
            string InputFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            string DictionaryFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + DictionaryFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilePath);
                Worksheet wsIF = wbIF.Worksheets[0];      //Input File First Worksheet
                int ifMaxRows = wsIF.Cells.MaxRow;
                int ifMaxColumns = wsIF.Cells.MaxColumn;

                #region Adding second worksheet to write Unmatched Noun-Modifier and Attributes
                if (wbIF.Worksheets.Count <= 1)
                    wbIF.Worksheets.Add();

                Worksheet wsUMA = wbIF.Worksheets[1];
                #endregion

                Workbook wbDF = new Workbook();
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbDF.Open(DictionaryFilePath);
                Worksheet wsDC = wbDF.Worksheets[0];      //Dictionary File First Worksheet
                int dMaxRows = wsDC.Cells.MaxRow;

                string ifNoun = string.Empty, ifModifier = string.Empty;
                bool NounModifierFoundInDictionary = false;
                bool UnmatchedAttributeFound = false;
                int daCol = 2;
                List<clsUMA> UnmatchedAttributeList = new List<clsUMA>();
                List<clsUNM> UnmatchedNounModifierList = new List<clsUNM>();

                #region Processing Rows starts
                for (int ifrow = 1; ifrow <= wsIF.Cells.MaxRow; ifrow++)
                {
                    ifNoun = wsIF.Cells[ifrow, 1].StringValue;
                    ifModifier = wsIF.Cells[ifrow, 2].StringValue;
                    NounModifierFoundInDictionary = false; UnmatchedAttributeFound = false;
                    for (int drow = 1; drow <= dMaxRows; drow++)
                    {
                        if (ifNoun.Trim().ToUpper() == wsDC.Cells[drow, 0].StringValue.Trim().ToUpper())
                        {
                            if (ifModifier.Trim().ToUpper() == wsDC.Cells[drow, 1].StringValue.Trim().ToUpper())
                            {
                                NounModifierFoundInDictionary = true;
                                daCol = 2;
                                for (int ifaCol = 3; ifaCol <= ifMaxColumns; ifaCol += 2)
                                {
                                    if (wsIF.Cells[ifrow, ifaCol].StringValue.Trim().ToUpper() != wsDC.Cells[drow, daCol].StringValue.Trim().ToUpper())
                                    {
                                        #region Add Dictionary row to Unmatched Attributes list
                                        if (UnmatchedAttributeList.Where(u => u.Source == "Dictionary" && u.Noun.Trim().ToUpper() == ifNoun.Trim().ToUpper() && u.Modifier.Trim().ToUpper() == ifModifier.Trim().ToUpper()).Count() == 0)
                                        {
                                            UnmatchedAttributeList.Add(new clsUMA()
                                            {
                                                Source = "Dictionary",
                                                Noun = wsDC.Cells[drow, 0].StringValue,
                                                Modifier = wsDC.Cells[drow, 1].StringValue,
                                                Attribute1 = wsDC.Cells[drow, 2].StringValue,
                                                Attribute2 = wsDC.Cells[drow, 3].StringValue,
                                                Attribute3 = wsDC.Cells[drow, 4].StringValue,
                                                Attribute4 = wsDC.Cells[drow, 5].StringValue,
                                                Attribute5 = wsDC.Cells[drow, 6].StringValue,
                                                Attribute6 = wsDC.Cells[drow, 7].StringValue,
                                                Attribute7 = wsDC.Cells[drow, 8].StringValue,
                                                Attribute8 = wsDC.Cells[drow, 9].StringValue,
                                                Attribute9 = wsDC.Cells[drow, 10].StringValue,
                                                Attribute10 = wsDC.Cells[drow, 11].StringValue,
                                                Attribute11 = wsDC.Cells[drow, 12].StringValue,
                                                Attribute12 = wsDC.Cells[drow, 13].StringValue,
                                                Attribute13 = wsDC.Cells[drow, 14].StringValue,
                                                Attribute14 = wsDC.Cells[drow, 15].StringValue,
                                                Attribute15 = wsDC.Cells[drow, 16].StringValue,
                                                Attribute16 = wsDC.Cells[drow, 17].StringValue,
                                                Attribute17 = wsDC.Cells[drow, 18].StringValue,
                                                Attribute18 = wsDC.Cells[drow, 19].StringValue,
                                                Attribute19 = wsDC.Cells[drow, 20].StringValue,
                                                Attribute20 = wsDC.Cells[drow, 21].StringValue,
                                                Attribute21 = wsDC.Cells[drow, 22].StringValue,
                                                Attribute22 = wsDC.Cells[drow, 23].StringValue,
                                                Attribute23 = wsDC.Cells[drow, 24].StringValue,
                                                Attribute24 = wsDC.Cells[drow, 25].StringValue,
                                                Attribute25 = wsDC.Cells[drow, 26].StringValue,
                                                Attribute26 = wsDC.Cells[drow, 27].StringValue,
                                                Attribute27 = wsDC.Cells[drow, 28].StringValue,
                                                Attribute28 = wsDC.Cells[drow, 29].StringValue,
                                                Attribute29 = wsDC.Cells[drow, 30].StringValue,
                                                Attribute30 = wsDC.Cells[drow, 31].StringValue,
                                                Attribute31 = wsDC.Cells[drow, 32].StringValue,
                                                Attribute32 = wsDC.Cells[drow, 33].StringValue,
                                                Attribute33 = wsDC.Cells[drow, 34].StringValue,
                                                Attribute34 = wsDC.Cells[drow, 35].StringValue,
                                                Attribute35 = wsDC.Cells[drow, 36].StringValue,
                                                Attribute36 = wsDC.Cells[drow, 37].StringValue,
                                                Attribute37 = wsDC.Cells[drow, 38].StringValue,
                                                Attribute38 = wsDC.Cells[drow, 39].StringValue,
                                                Attribute39 = wsDC.Cells[drow, 40].StringValue,
                                                Attribute40 = wsDC.Cells[drow, 41].StringValue,
                                                Attribute41 = wsDC.Cells[drow, 42].StringValue,
                                                Attribute42 = wsDC.Cells[drow, 43].StringValue,
                                                Attribute43 = wsDC.Cells[drow, 44].StringValue,
                                                Attribute44 = wsDC.Cells[drow, 45].StringValue,
                                                Attribute45 = wsDC.Cells[drow, 46].StringValue,
                                                Attribute46 = wsDC.Cells[drow, 47].StringValue,
                                                Attribute47 = wsDC.Cells[drow, 48].StringValue,
                                                Attribute48 = wsDC.Cells[drow, 49].StringValue,
                                                Attribute49 = wsDC.Cells[drow, 50].StringValue,
                                                Attribute50 = wsDC.Cells[drow, 51].StringValue
                                            });
                                        }
                                        #endregion

                                        #region Add Input file row to Unmatched Attributes list
                                        UnmatchedAttributeList.Add(new clsUMA()
                                        {
                                            Source = "Input",
                                            Noun = wsIF.Cells[ifrow, 1].StringValue,
                                            Modifier = wsIF.Cells[ifrow, 2].StringValue,
                                            Attribute1 = wsIF.Cells[ifrow, 3].StringValue,
                                            Attribute2 = wsIF.Cells[ifrow, 5].StringValue,
                                            Attribute3 = wsIF.Cells[ifrow, 7].StringValue,
                                            Attribute4 = wsIF.Cells[ifrow, 9].StringValue,
                                            Attribute5 = wsIF.Cells[ifrow, 11].StringValue,
                                            Attribute6 = wsIF.Cells[ifrow, 13].StringValue,
                                            Attribute7 = wsIF.Cells[ifrow, 15].StringValue,
                                            Attribute8 = wsIF.Cells[ifrow, 17].StringValue,
                                            Attribute9 = wsIF.Cells[ifrow, 19].StringValue,
                                            Attribute10 = wsIF.Cells[ifrow, 21].StringValue,
                                            Attribute11 = wsIF.Cells[ifrow, 23].StringValue,
                                            Attribute12 = wsIF.Cells[ifrow, 25].StringValue,
                                            Attribute13 = wsIF.Cells[ifrow, 27].StringValue,
                                            Attribute14 = wsIF.Cells[ifrow, 29].StringValue,
                                            Attribute15 = wsIF.Cells[ifrow, 31].StringValue,
                                            Attribute16 = wsIF.Cells[ifrow, 33].StringValue,
                                            Attribute17 = wsIF.Cells[ifrow, 35].StringValue,
                                            Attribute18 = wsIF.Cells[ifrow, 37].StringValue,
                                            Attribute19 = wsIF.Cells[ifrow, 39].StringValue,
                                            Attribute20 = wsIF.Cells[ifrow, 41].StringValue,
                                            Attribute21 = wsIF.Cells[ifrow, 43].StringValue,
                                            Attribute22 = wsIF.Cells[ifrow, 45].StringValue,
                                            Attribute23 = wsIF.Cells[ifrow, 47].StringValue,
                                            Attribute24 = wsIF.Cells[ifrow, 49].StringValue,
                                            Attribute25 = wsIF.Cells[ifrow, 51].StringValue,
                                            Attribute26 = wsIF.Cells[ifrow, 53].StringValue,
                                            Attribute27 = wsIF.Cells[ifrow, 55].StringValue,
                                            Attribute28 = wsIF.Cells[ifrow, 57].StringValue,
                                            Attribute29 = wsIF.Cells[ifrow, 59].StringValue,
                                            Attribute30 = wsIF.Cells[ifrow, 61].StringValue,
                                            Attribute31 = wsIF.Cells[ifrow, 63].StringValue,
                                            Attribute32 = wsIF.Cells[ifrow, 65].StringValue,
                                            Attribute33 = wsIF.Cells[ifrow, 67].StringValue,
                                            Attribute34 = wsIF.Cells[ifrow, 69].StringValue,
                                            Attribute35 = wsIF.Cells[ifrow, 71].StringValue,
                                            Attribute36 = wsIF.Cells[ifrow, 73].StringValue,
                                            Attribute37 = wsIF.Cells[ifrow, 75].StringValue,
                                            Attribute38 = wsIF.Cells[ifrow, 77].StringValue,
                                            Attribute39 = wsIF.Cells[ifrow, 79].StringValue,
                                            Attribute40 = wsIF.Cells[ifrow, 81].StringValue,
                                            Attribute41 = wsIF.Cells[ifrow, 83].StringValue,
                                            Attribute42 = wsIF.Cells[ifrow, 85].StringValue,
                                            Attribute43 = wsIF.Cells[ifrow, 87].StringValue,
                                            Attribute44 = wsIF.Cells[ifrow, 89].StringValue,
                                            Attribute45 = wsIF.Cells[ifrow, 91].StringValue,
                                            Attribute46 = wsIF.Cells[ifrow, 93].StringValue,
                                            Attribute47 = wsIF.Cells[ifrow, 95].StringValue,
                                            Attribute48 = wsIF.Cells[ifrow, 97].StringValue,
                                            Attribute49 = wsIF.Cells[ifrow, 99].StringValue,
                                            Attribute50 = wsIF.Cells[ifrow, 101].StringValue
                                        });
                                        #endregion

                                        UnmatchedAttributeFound = true;
                                        break;
                                    }
                                    daCol++;
                                }
                                if (UnmatchedAttributeFound)
                                    break;
                            }
                        }
                    }

                    if (!NounModifierFoundInDictionary)
                    {
                        if (UnmatchedNounModifierList.Where(u => u.Noun.Trim().ToUpper() == ifNoun.Trim().ToUpper() && u.Modifier.Trim().ToUpper() == ifModifier.Trim().ToUpper()).Count() == 0)
                        {
                            UnmatchedNounModifierList.Add(new clsUNM()
                            {
                                Noun = ifNoun,
                                Modifier = ifModifier
                            });
                        }
                    }
                }
                #endregion

                #region Writing Unmatched Noun Modifiers to input file second worksheet
                wsUMA.Cells[0, 0].PutValue("Noun");
                wsUMA.Cells[0, 1].PutValue("Modifier");
                int wsUMNMLastRowNo = 0;
                for (int idx = 0; idx < UnmatchedNounModifierList.Count(); idx++)
                {
                    wsUMNMLastRowNo++;
                    wsUMA.Cells[wsUMNMLastRowNo, 0].PutValue(UnmatchedNounModifierList[idx].Noun);
                    wsUMA.Cells[wsUMNMLastRowNo, 1].PutValue(UnmatchedNounModifierList[idx].Modifier);
                }
                #endregion

                #region Write headers to input file second worksheet
                int wsUMALastRowNo = wsUMA.Cells.MaxRow + 1;

                wsUMA.Cells[wsUMALastRowNo, 0].PutValue("Source");
                wsUMA.Cells[wsUMALastRowNo, 1].PutValue("Noun");
                wsUMA.Cells[wsUMALastRowNo, 2].PutValue("Modifier");
                int wsUMAcolNo = 3;
                for (int aCtr = 1; aCtr <= 50; aCtr++)
                {
                    wsUMA.Cells[wsUMALastRowNo, wsUMAcolNo].PutValue("Attribute" + aCtr.ToString());
                    wsUMAcolNo++;
                }
                #endregion

                #region Writing Unmatched Attributes to input file second worksheet
                bool IsAttributeFound = false;
                wsUMNMLastRowNo += 2;
                for (int i = 0; i < UnmatchedAttributeList.Count(); i++)
                {
                    IsAttributeFound = false;
                    for (int r = wsUMNMLastRowNo; r <= wsUMA.Cells.MaxRow; r++)
                    {
                        if (wsUMA.Cells[r, 0].StringValue.Trim() == "Input" &&
                            wsUMA.Cells[r, 1].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Noun.Trim().ToUpper() &&
                            wsUMA.Cells[r, 2].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Modifier.Trim().ToUpper() &&
                            wsUMA.Cells[r, 3].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute1.Trim().ToUpper() &&
                            wsUMA.Cells[r, 4].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute2.Trim().ToUpper() &&
                            wsUMA.Cells[r, 5].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute3.Trim().ToUpper() &&
                            wsUMA.Cells[r, 6].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute4.Trim().ToUpper() &&
                            wsUMA.Cells[r, 7].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute5.Trim().ToUpper() &&
                            wsUMA.Cells[r, 8].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute6.Trim().ToUpper() &&
                            wsUMA.Cells[r, 9].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute7.Trim().ToUpper() &&
                            wsUMA.Cells[r, 10].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute8.Trim().ToUpper() &&
                            wsUMA.Cells[r, 11].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute9.Trim().ToUpper() &&
                            wsUMA.Cells[r, 12].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute10.Trim().ToUpper() &&
                            wsUMA.Cells[r, 13].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute11.Trim().ToUpper() &&
                            wsUMA.Cells[r, 14].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute12.Trim().ToUpper() &&
                            wsUMA.Cells[r, 15].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute13.Trim().ToUpper() &&
                            wsUMA.Cells[r, 16].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute14.Trim().ToUpper() &&
                            wsUMA.Cells[r, 17].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute15.Trim().ToUpper() &&
                            wsUMA.Cells[r, 18].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute16.Trim().ToUpper() &&
                            wsUMA.Cells[r, 19].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute17.Trim().ToUpper() &&
                            wsUMA.Cells[r, 20].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute18.Trim().ToUpper() &&
                            wsUMA.Cells[r, 21].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute19.Trim().ToUpper() &&
                            wsUMA.Cells[r, 22].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute20.Trim().ToUpper() &&
                            wsUMA.Cells[r, 23].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute21.Trim().ToUpper() &&
                            wsUMA.Cells[r, 24].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute22.Trim().ToUpper() &&
                            wsUMA.Cells[r, 25].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute23.Trim().ToUpper() &&
                            wsUMA.Cells[r, 26].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute24.Trim().ToUpper() &&
                            wsUMA.Cells[r, 27].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute25.Trim().ToUpper() &&
                            wsUMA.Cells[r, 28].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute26.Trim().ToUpper() &&
                            wsUMA.Cells[r, 29].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute27.Trim().ToUpper() &&
                            wsUMA.Cells[r, 30].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute28.Trim().ToUpper() &&
                            wsUMA.Cells[r, 31].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute29.Trim().ToUpper() &&
                            wsUMA.Cells[r, 32].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute30.Trim().ToUpper() &&
                            wsUMA.Cells[r, 33].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute31.Trim().ToUpper() &&
                            wsUMA.Cells[r, 34].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute32.Trim().ToUpper() &&
                            wsUMA.Cells[r, 35].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute33.Trim().ToUpper() &&
                            wsUMA.Cells[r, 36].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute34.Trim().ToUpper() &&
                            wsUMA.Cells[r, 37].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute35.Trim().ToUpper() &&
                            wsUMA.Cells[r, 38].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute36.Trim().ToUpper() &&
                            wsUMA.Cells[r, 39].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute37.Trim().ToUpper() &&
                            wsUMA.Cells[r, 40].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute38.Trim().ToUpper() &&
                            wsUMA.Cells[r, 41].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute39.Trim().ToUpper() &&
                            wsUMA.Cells[r, 42].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute40.Trim().ToUpper() &&
                            wsUMA.Cells[r, 43].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute41.Trim().ToUpper() &&
                            wsUMA.Cells[r, 44].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute42.Trim().ToUpper() &&
                            wsUMA.Cells[r, 45].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute43.Trim().ToUpper() &&
                            wsUMA.Cells[r, 46].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute44.Trim().ToUpper() &&
                            wsUMA.Cells[r, 47].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute45.Trim().ToUpper() &&
                            wsUMA.Cells[r, 48].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute46.Trim().ToUpper() &&
                            wsUMA.Cells[r, 49].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute47.Trim().ToUpper() &&
                            wsUMA.Cells[r, 50].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute48.Trim().ToUpper() &&
                            wsUMA.Cells[r, 51].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute49.Trim().ToUpper() &&
                            wsUMA.Cells[r, 52].StringValue.Trim().ToUpper() == UnmatchedAttributeList[i].Attribute50.Trim().ToUpper())
                        {
                            IsAttributeFound = true;
                            break;
                        }
                    }

                    if (!IsAttributeFound)
                    {
                        wsUMALastRowNo++;
                        wsUMA.Cells[wsUMALastRowNo, 0].PutValue(UnmatchedAttributeList[i].Source);
                        wsUMA.Cells[wsUMALastRowNo, 1].PutValue(UnmatchedAttributeList[i].Noun);
                        wsUMA.Cells[wsUMALastRowNo, 2].PutValue(UnmatchedAttributeList[i].Modifier);
                        wsUMA.Cells[wsUMALastRowNo, 3].PutValue(UnmatchedAttributeList[i].Attribute1);
                        wsUMA.Cells[wsUMALastRowNo, 4].PutValue(UnmatchedAttributeList[i].Attribute2);
                        wsUMA.Cells[wsUMALastRowNo, 5].PutValue(UnmatchedAttributeList[i].Attribute3);
                        wsUMA.Cells[wsUMALastRowNo, 6].PutValue(UnmatchedAttributeList[i].Attribute4);
                        wsUMA.Cells[wsUMALastRowNo, 7].PutValue(UnmatchedAttributeList[i].Attribute5);
                        wsUMA.Cells[wsUMALastRowNo, 8].PutValue(UnmatchedAttributeList[i].Attribute6);
                        wsUMA.Cells[wsUMALastRowNo, 9].PutValue(UnmatchedAttributeList[i].Attribute7);
                        wsUMA.Cells[wsUMALastRowNo, 10].PutValue(UnmatchedAttributeList[i].Attribute8);
                        wsUMA.Cells[wsUMALastRowNo, 11].PutValue(UnmatchedAttributeList[i].Attribute9);
                        wsUMA.Cells[wsUMALastRowNo, 12].PutValue(UnmatchedAttributeList[i].Attribute10);
                        wsUMA.Cells[wsUMALastRowNo, 13].PutValue(UnmatchedAttributeList[i].Attribute11);
                        wsUMA.Cells[wsUMALastRowNo, 14].PutValue(UnmatchedAttributeList[i].Attribute12);
                        wsUMA.Cells[wsUMALastRowNo, 15].PutValue(UnmatchedAttributeList[i].Attribute13);
                        wsUMA.Cells[wsUMALastRowNo, 16].PutValue(UnmatchedAttributeList[i].Attribute14);
                        wsUMA.Cells[wsUMALastRowNo, 17].PutValue(UnmatchedAttributeList[i].Attribute15);
                        wsUMA.Cells[wsUMALastRowNo, 18].PutValue(UnmatchedAttributeList[i].Attribute16);
                        wsUMA.Cells[wsUMALastRowNo, 19].PutValue(UnmatchedAttributeList[i].Attribute17);
                        wsUMA.Cells[wsUMALastRowNo, 20].PutValue(UnmatchedAttributeList[i].Attribute18);
                        wsUMA.Cells[wsUMALastRowNo, 21].PutValue(UnmatchedAttributeList[i].Attribute19);
                        wsUMA.Cells[wsUMALastRowNo, 22].PutValue(UnmatchedAttributeList[i].Attribute20);
                        wsUMA.Cells[wsUMALastRowNo, 23].PutValue(UnmatchedAttributeList[i].Attribute21);
                        wsUMA.Cells[wsUMALastRowNo, 24].PutValue(UnmatchedAttributeList[i].Attribute22);
                        wsUMA.Cells[wsUMALastRowNo, 25].PutValue(UnmatchedAttributeList[i].Attribute23);
                        wsUMA.Cells[wsUMALastRowNo, 26].PutValue(UnmatchedAttributeList[i].Attribute24);
                        wsUMA.Cells[wsUMALastRowNo, 27].PutValue(UnmatchedAttributeList[i].Attribute25);
                        wsUMA.Cells[wsUMALastRowNo, 28].PutValue(UnmatchedAttributeList[i].Attribute26);
                        wsUMA.Cells[wsUMALastRowNo, 29].PutValue(UnmatchedAttributeList[i].Attribute27);
                        wsUMA.Cells[wsUMALastRowNo, 30].PutValue(UnmatchedAttributeList[i].Attribute28);
                        wsUMA.Cells[wsUMALastRowNo, 31].PutValue(UnmatchedAttributeList[i].Attribute29);
                        wsUMA.Cells[wsUMALastRowNo, 32].PutValue(UnmatchedAttributeList[i].Attribute30);
                        wsUMA.Cells[wsUMALastRowNo, 33].PutValue(UnmatchedAttributeList[i].Attribute31);
                        wsUMA.Cells[wsUMALastRowNo, 34].PutValue(UnmatchedAttributeList[i].Attribute32);
                        wsUMA.Cells[wsUMALastRowNo, 35].PutValue(UnmatchedAttributeList[i].Attribute33);
                        wsUMA.Cells[wsUMALastRowNo, 36].PutValue(UnmatchedAttributeList[i].Attribute34);
                        wsUMA.Cells[wsUMALastRowNo, 37].PutValue(UnmatchedAttributeList[i].Attribute35);
                        wsUMA.Cells[wsUMALastRowNo, 38].PutValue(UnmatchedAttributeList[i].Attribute36);
                        wsUMA.Cells[wsUMALastRowNo, 39].PutValue(UnmatchedAttributeList[i].Attribute37);
                        wsUMA.Cells[wsUMALastRowNo, 40].PutValue(UnmatchedAttributeList[i].Attribute38);
                        wsUMA.Cells[wsUMALastRowNo, 41].PutValue(UnmatchedAttributeList[i].Attribute39);
                        wsUMA.Cells[wsUMALastRowNo, 42].PutValue(UnmatchedAttributeList[i].Attribute40);
                        wsUMA.Cells[wsUMALastRowNo, 43].PutValue(UnmatchedAttributeList[i].Attribute41);
                        wsUMA.Cells[wsUMALastRowNo, 44].PutValue(UnmatchedAttributeList[i].Attribute42);
                        wsUMA.Cells[wsUMALastRowNo, 45].PutValue(UnmatchedAttributeList[i].Attribute43);
                        wsUMA.Cells[wsUMALastRowNo, 46].PutValue(UnmatchedAttributeList[i].Attribute44);
                        wsUMA.Cells[wsUMALastRowNo, 47].PutValue(UnmatchedAttributeList[i].Attribute45);
                        wsUMA.Cells[wsUMALastRowNo, 48].PutValue(UnmatchedAttributeList[i].Attribute46);
                        wsUMA.Cells[wsUMALastRowNo, 49].PutValue(UnmatchedAttributeList[i].Attribute47);
                        wsUMA.Cells[wsUMALastRowNo, 50].PutValue(UnmatchedAttributeList[i].Attribute48);
                        wsUMA.Cells[wsUMALastRowNo, 51].PutValue(UnmatchedAttributeList[i].Attribute49);
                        wsUMA.Cells[wsUMALastRowNo, 52].PutValue(UnmatchedAttributeList[i].Attribute50);
                    }
                }
                #endregion

                #region Save and Download the file
                wsUMA.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

                if (File.Exists(InputFilePath))
                    File.Delete(InputFilePath);

                if (File.Exists(DictionaryFilePath))
                    File.Delete(DictionaryFilePath);

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

        #region Download Help Document
        [HttpGet]
        [Route("DownloadHelpDocument")]
        public HttpResponseMessage DownloadHelpDocument()
        {
            string HelpFilepath = HttpContext.Current.Server.MapPath(@"\HelpDocs\DictionaryCheck.pptx");

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Check whether File exists.
                if (!File.Exists(HelpFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(HelpFilepath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = HelpFilepath;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(HelpFilepath));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion
    }
}
