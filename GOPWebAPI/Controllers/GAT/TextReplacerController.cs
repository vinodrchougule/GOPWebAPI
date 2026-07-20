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
    [RoutePrefix("api/TextReplacer")]
    public class TextReplacerController : ApiController
    {
        private BLLGAT _BLLGAT;
        private BLLGATTextReplacer _BLLGATTextReplacer;

        public TextReplacerController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
            _BLLGATTextReplacer = new BLLGATTextReplacer(connectionString);
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

                if(wbIF.Worksheets.Count < 2)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File should have at least 2 worksheets.");
                }

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Name.ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First worksheet with name 'Description' not found in input file. Please select a valid file.");
                }
                
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Sl No' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "ITEM ID")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Item ID' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "INPUT DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Input Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "DESCRIPTION1")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Description1' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 4].StringValue.Trim().ToUpper() != "DESCRIPTION2")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Description2' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating the second worksheet
                Worksheet ws2 = wbIF.Worksheets[1];
                if (ws2.Name.ToUpper() != "TERMS")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second worksheet with name 'Terms' not found in input file. Please select a valid file.");
                }

                if (ws2.Cells[0, 0].StringValue.Trim().ToUpper() != "INDEX")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Index' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 1].StringValue.Trim().ToUpper() != "TERM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Term' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 2].StringValue.Trim().ToUpper() != "FROM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'From' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 3].StringValue.Trim().ToUpper() != "TO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'To' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells[0, 4].StringValue.Trim().ToUpper() != "NO OF FINDINGS")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'No Of Findings' not found in input file second worksheet. Please select a valid file.");
                }

                if (ws2.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file second worksheet has no Terms to replace.");
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

        #region Write Description and terms input data to database tables
        [HttpPost]
        [Route("WriteDescriptionAndTermsInputDataToDatabase")]
        public HttpResponseMessage WriteDescriptionAndTermsInputDataToDatabase(string InputFileName)
        {
            string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            string sqlTableName1, sqlTableName2 = string.Empty;
            try
            {
                #region Create temp tables and write uploaded files data
                sqlTableName1 = _BLLGAT.CreateSQLTable("spGATTRCreateDescTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(InputFileName, sqlTableName1, 0);
                if (IsSucceeded)
                {
                    sqlTableName2 = _BLLGAT.CreateSQLTable("spGATTRCreateTermsTempTable");
                    IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(InputFileName, sqlTableName2, 1);
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName1 + "," + sqlTableName2);
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

        #region Replace Text and Write the Output to file
        [HttpPost]
        [Route("ReplaceTextAndWriteOutputToExcel")]
        public HttpResponseMessage ReplaceTextAndWriteOutputToExcel(string InputFileName, string UploadedInputFileName, string descSqlTableName, string termsSqlTableName, string Delimiter = "Space")
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);
                Worksheet sws1 = wbIF.Worksheets[0];
                Worksheet sws2 = wbIF.Worksheets[1];

                #region Replace Description And Get Data
                DataTable dataTable = _BLLGATTextReplacer.ReplaceDescriptionAndGetData(descSqlTableName, termsSqlTableName, Delimiter);
                #endregion

                #region Get the terms with count
                DataTable dataTable1 = _BLLGATTextReplacer.GetTermsWithCount(descSqlTableName, termsSqlTableName, Delimiter);
                #endregion

                int RowCounter = 1;
                if (Delimiter.Trim().ToUpper() == "COMMA & SPACE")
                {
                    #region Write Description back to input file
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        sws1.Cells[i + 1, 3].PutValue(dataTable.Rows[i]["Description1"].ToString());
                        sws1.Cells[i + 1, 4].PutValue(dataTable.Rows[i]["Description2"].ToString());
                    }
                    #endregion
                }
                else
                {
                    string ItemID = "", Word = "", pItemID = "", Description1 = "";
                    int iCnt = 1, wrCnt = 1;

                    #region Write Description back to input file
                    foreach (DataRow dr in dataTable.Rows)
                    {
                        ItemID = dr["ItemID"].ToString();
                        Word = dr["Word"].ToString();

                        if (iCnt == 1)
                        {
                            Description1 = Word;
                            iCnt++;
                        }
                        else if (ItemID == pItemID)
                            Description1 = Description1 + Delimiter + Word;
                        else
                        {
                            sws1.Cells[wrCnt, 3].PutValue(Description1);
                            Description1 = Word;
                            wrCnt++;
                        }
                        pItemID = ItemID;
                    }
                    sws1.Cells[wrCnt, 3].PutValue(Description1);
                    #endregion
                }

                #region Write Terms Count Data back to input file
                RowCounter = 1;
                foreach (DataRow dr in dataTable1.Rows)
                {
                    sws2.Cells[RowCounter, 0].PutValue(dr["Index"].ToString());
                    sws2.Cells[RowCounter, 1].PutValue(dr["Term"].ToString());
                    sws2.Cells[RowCounter, 2].PutValue(dr["From"].ToString());
                    sws2.Cells[RowCounter, 3].PutValue(dr["To"].ToString());
                    sws2.Cells[RowCounter, 4].PutValue(Convert.ToInt32(dr["NoOfFindings"]));
                    RowCounter++;
                }
                #endregion

                sws1.AutoFitColumns();
                sws2.AutoFitColumns();
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
