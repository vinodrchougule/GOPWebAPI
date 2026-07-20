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

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/TermAnalysis")]
    public class TermAnalysisController : ApiController
    {
        private BLLGAT _BLLGAT;
        private BLLGATTermAnalysis _BLLGATTermAnalysis;
        public TermAnalysisController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
            _BLLGATTermAnalysis = new BLLGATTermAnalysis(connectionString);
        }

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

                if (wbIF.Worksheets.Count < 4)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File should have at least 4 worksheets.");
                }

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Name.ToUpper() != "INPUT")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First worksheet with name 'Input' not found in input file. Please select a valid file.");
                }

                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "ITEM ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Item ID' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                if (ws2.Name.Trim().ToUpper() != "TERM FREQUENCY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second worksheet with name 'Term Frequency' not found in input file. Please select a valid file.");
                }

                if (ws2.Cells[0, 0].StringValue.Trim().ToUpper() != "TERM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Term' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 1].StringValue.Trim().ToUpper() != "FREQUENCY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Frequency' not found in input file second worksheet. Please select a valid file.");
                }
                #endregion

                #region Validating third worksheet
                Worksheet ws3 = wbIF.Worksheets[2];
                if (ws3.Name.Trim().ToUpper() != "ATTRIBUTES")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third worksheet with name 'Attributes' not found in input file. Please select a valid file.");
                }

                if (ws3.Cells[0, 0].StringValue.Trim().ToUpper() != "ITEM ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Item ID' not found in input file third worksheet. Please select a valid file.");
                }

                if (ws3.Cells[0, 1].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Description' not found in input file third worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCtr = 1; aCtr <= 50; aCtr++)
                {
                    hdr = ws3.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' in input file third worksheet.Please download format from Template.");
                    }
                    attNo++;
                }
                #endregion

                #region Validating fourth worksheet
                Worksheet ws4 = wbIF.Worksheets[3];
                if (ws4.Name.ToUpper() != "REPEATED WORDS")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth worksheet with name 'Repeated Words' not found in input file. Please select a valid file.");
                }

                if (ws4.Cells[0, 0].StringValue.Trim().ToUpper() != "ITEM ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Item ID' not found in input file fourth worksheet. Please select a valid file.");
                }

                if (ws4.Cells[0, 1].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Description' not found in input file fourth worksheet. Please select a valid file.");
                }

                attNo = 1;
                for (int aCtr = 1; aCtr <= 20; aCtr += 2)
                {
                    hdr = ws4.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();

                    if (hdr != "REPEATED WORD" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Repeated Word" + attNo.ToString() + "' in fourth worksheet.Please download format from Template.");
                    }

                    hdr = ws4.Cells[0, aCtr + 2].StringValue.Trim().ToUpper();
                    if (hdr != "REPEATED WORD" + attNo.ToString() + " - ATTRIBUTE NUMBERS")
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 3).ToString() + "] must contain 'Repeated Word" + attNo.ToString() + " - Attribute Numbers' in fourth worksheet.Please download format from Template.");
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

        #region Write input file first worksheet description data to Database
        [HttpPost]
        [Route("WriteInputDescriptionDataToDatabase")]
        public HttpResponseMessage WriteInputDescriptionDataToDatabase(string InputFileName)
        {
            string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            string sqlTableName = string.Empty;

            try
            {
                #region Create temp tables and write uploaded file data
                sqlTableName = _BLLGAT.CreateSQLTable("spGATTACreateTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(InputFileName, sqlTableName, 0);
                #endregion

                if(IsSucceeded)
                    return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Writing data to server failed");
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

        #region Analyze Terms and Write the Output to file
        [HttpPost]
        [Route("AnalyzeTermsAndWriteOutputToExcel")]
        public HttpResponseMessage AnalyzeTermsAndWriteOutputToExcel(string InputFileName, string UploadedInputFileName, string descSqlTableName,string Delimiter = " ")
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                char actualDelimiter = string.IsNullOrEmpty(Delimiter) ? ' ' : Delimiter[0];

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);

                #region Write Term Frequency
                Worksheet wsTF = wbIF.Worksheets[1];
                DataTable dtTermFrequency = _BLLGATTermAnalysis.GetTermFrequencyData(descSqlTableName, actualDelimiter);

                Aspose.Cells.Style styleCenter = wsTF.Cells[0, 3].GetStyle();

                styleCenter.IsTextWrapped = true;
                styleCenter.HorizontalAlignment = TextAlignmentType.Center;
                styleCenter.VerticalAlignment = TextAlignmentType.Center;
                styleCenter.Font.Size = 10;

                Aspose.Cells.Style styleLeft = wsTF.Cells[0, 3].GetStyle();
                styleLeft.IsTextWrapped = true;
                styleLeft.HorizontalAlignment = TextAlignmentType.Left;
                styleLeft.VerticalAlignment = TextAlignmentType.Center;
                styleLeft.Font.Size = 10;

                #region Write Term Frequency back to input file
                for (int i = 0; i < dtTermFrequency.Rows.Count; i++)
                {
                    wsTF.Cells[i + 1, 0].PutValue(dtTermFrequency.Rows[i]["Term"].ToString());
                    wsTF.Cells[i + 1, 0].SetStyle(styleLeft);
                    if(dtTermFrequency.Rows[i]["Frequency"] == DBNull.Value)
                        wsTF.Cells[i + 1, 1].PutValue(0);
                    else
                        wsTF.Cells[i + 1, 1].PutValue(Convert.ToInt32(dtTermFrequency.Rows[i]["Frequency"]));
                    wsTF.Cells[i + 1, 0].SetStyle(styleCenter);
                }
                #endregion
                #endregion

                #region Write Term Attributes
                Worksheet wsTA = wbIF.Worksheets[2];
                DataTable dtTermAttributes = _BLLGATTermAnalysis.GetTermAttributesData(descSqlTableName);

                #region Write Term Attributes back to input file
                for (int i = 0; i < dtTermAttributes.Rows.Count; i++)
                {
                    wsTA.Cells[i + 1, 0].PutValue(dtTermAttributes.Rows[i]["ItemID"].ToString());
                    wsTA.Cells[i + 1, 1].PutValue(dtTermAttributes.Rows[i]["Description"].ToString());
                    string Description = dtTermAttributes.Rows[i]["Description"].ToString();
                    if (!string.IsNullOrEmpty(Description))
                    {
                        string[] attrArray = Description.Split(actualDelimiter);
                        int c = 2;
                        for (int a = 0; a < attrArray.Count(); a++)
                        {
                            if (attrArray[a].Trim().Length > 0)
                            {
                                wsTA.Cells[i + 1, c].PutValue(attrArray[a].Trim());
                                c++;
                            }
                        }
                    }
                }
                #endregion
                #endregion

                #region Write Repeated Words
                Worksheet wsRW = wbIF.Worksheets[3];
                DataTable dtRepeatedWords = _BLLGATTermAnalysis.GetRepeatedWordsData(descSqlTableName);

                for (int i = 0; i < dtRepeatedWords.Rows.Count; i++)
                {
                    wsRW.Cells[i + 1, 0].PutValue(dtTermAttributes.Rows[i]["ItemID"].ToString());
                    wsRW.Cells[i + 1, 1].PutValue(dtTermAttributes.Rows[i]["Description"].ToString());
                    string Description = dtTermAttributes.Rows[i]["Description"].ToString().Trim();
                    string word = "", RepeatedWord = "", AttributeNos = "";
                    if (!string.IsNullOrEmpty(Description))
                    {
                        string[] arr = Description.Split(actualDelimiter);
                        for (int j = 0; j < arr.Count(); j++)
                        {
                            word = arr[j].Trim();
                            int ctr = 0;
                            for (int k = 0; k < arr.Count(); k++)
                            {
                                if (word.ToUpper() == arr[k].Trim().ToUpper())
                                    ctr++;
                            }

                            if (ctr > 1)
                            {
                                RepeatedWord = word;
                                Boolean IsRepeatedWordPrinted = false;
                                int BlankColNo = 0;
                                for (int col = 2; col <= 21; col += 2)
                                {
                                    if (string.IsNullOrEmpty(wsRW.Cells[i + 1, col].StringValue))
                                    {
                                        BlankColNo = col;
                                        break;
                                    }
                                    else
                                    {
                                        if (wsRW.Cells[i + 1, col].StringValue.ToString().Trim().ToUpper() == RepeatedWord.Trim().ToUpper())
                                            IsRepeatedWordPrinted = true;
                                    }
                                }

                                if (!IsRepeatedWordPrinted)
                                {
                                    AttributeNos = "";
                                    for (int m = 0; m < arr.Count(); m++)
                                    {
                                        if (RepeatedWord.ToUpper() == arr[m].Trim().ToUpper())
                                        {
                                            if (AttributeNos.Trim().Length == 0)
                                                AttributeNos = (m + 1).ToString();
                                            else
                                                AttributeNos = AttributeNos + "," + (m + 1).ToString();
                                        }
                                    }

                                    wsRW.Cells[i + 1, BlankColNo].PutValue(RepeatedWord.Trim());
                                    wsRW.Cells[i + 1, BlankColNo + 1].PutValue(AttributeNos.Trim());
                                }
                            }
                        }
                    }
                }
                #endregion

                wsTF.AutoFitColumns();
                wsTA.AutoFitColumns();
                wsRW.AutoFitColumns();
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
