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
using System.Xml.Linq;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/STPInputFormatConversion")]
    public class STPInputFormatConversionController : ApiController
    {
        private BLLGAT _BLLGAT;
        private BLLGATSTPInputFormatConversion _BLLGATSTPInputFormatConversion;
        public STPInputFormatConversionController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
            _BLLGATSTPInputFormatConversion = new BLLGATSTPInputFormatConversion(connectionString);
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

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "SLNO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'slno' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "MATERIAL")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Material' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "PO TEXT")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'PO Text' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 4].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [E1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 5].StringValue.Trim().ToUpper() != "MNFR")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [F1] with Value 'MNFR' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 6].StringValue.Trim().ToUpper() != "MNFR ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [G1] with Value 'MNFR Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 7].StringValue.Trim().ToUpper() != "PTNO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [H1] with Value 'PTNO' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 8].StringValue.Trim().ToUpper() != "PTNO ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [I1] with Value 'PTNO Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 9].StringValue.Trim().ToUpper() != "MNFR 1")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [J1] with Value 'MNFR 1' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 10].StringValue.Trim().ToUpper() != "MNFR 1 ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [K1] with Value 'MNFR 1 Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 11].StringValue.Trim().ToUpper() != "PTNO 1")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [L1] with Value 'PTNO 1' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 12].StringValue.Trim().ToUpper() != "PTNO 1 ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [M1] with Value 'PTNO 1 Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 13].StringValue.Trim().ToUpper() != "MNFR 2")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [N1] with Value 'MNFR 2' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 14].StringValue.Trim().ToUpper() != "MNFR 2 ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [O1] with Value 'MNFR 2 Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 15].StringValue.Trim().ToUpper() != "PTNO 2")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [P1] with Value 'PTNO 2' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 16].StringValue.Trim().ToUpper() != "PTNO 2 ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [Q1] with Value 'PTNO 2 Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 17].StringValue.Trim().ToUpper() != "DWG")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [R1] with Value 'DWG' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 18].StringValue.Trim().ToUpper() != "DWG ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [S1] with Value 'DWG Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 19].StringValue.Trim().ToUpper() != "DRAWING VERSION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [T1] with Value 'Drawing Version' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 20].StringValue.Trim().ToUpper() != "DRAWING VERSION ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [U1] with Value 'Drawing Version Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 21].StringValue.Trim().ToUpper() != "POS/TAG")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [V1] with Value 'POS/TAG' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 22].StringValue.Trim().ToUpper() != "POS/TAG ATTRIBUTE VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [W1] with Value 'POS/TAG Attribute Value' not found in input file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCtr = 23; aCtr <= 93; aCtr += 2)
                {
                    hdr = ws1.Cells[0, aCtr].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE NAME " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 1).ToString() + "] must contain 'Attribute Name " + attNo.ToString() + "' .Please download format from Template.");
                    }

                    hdr = ws1.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Attribute Value " + attNo.ToString() + "' .Please download format from Template.");
                    }
                    attNo++;
                }

                if (ws1.Cells[0, 95].StringValue.Trim().ToUpper() != "REMAINING_INFORMATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [CR1] with Value 'Remaining_Information' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
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

        #region Validate the taxonomy uploaded file
        [HttpPost]
        [Route("ValidateTaxonomyFile")]
        public HttpResponseMessage ValidateTaxonomyFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid taxonomy file of xlsx format only");

                Workbook wbTF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbTF.Open(UploadedFilepath);

                #region Validating the first worksheet
                Worksheet ws1 = wbTF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'sl no' not found in taxonomy file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Reference' not found in taxonomy file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Noun' not found in taxonomy file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 3].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [D1] with Value 'Modifier' not found in taxonomy file first worksheet. Please select a valid file.");
                }

                int attNo = 1;
                string hdr;
                for (int aCtr = 4; aCtr <= 74; aCtr += 2)
                {
                    hdr = ws1.Cells[0, aCtr].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE NAME" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 1).ToString() + "] must contain 'Attribute Name" + attNo.ToString() + "' .Please download format from taxonomy Template.");
                    }

                    hdr = ws1.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Attribute Value" + attNo.ToString() + "' .Please download format from taxonomy Template.");
                    }
                    attNo++;
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Taxonomy file first Worksheet has no data rows.");
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

        #region Write input and taxonomy files data to database tables
        [HttpPost]
        [Route("WriteInputAndTaxonomyDataToDatabase")]
        public HttpResponseMessage WriteInputAndTaxonomyDataToDatabase(string InputFileName, string TaxonomyFileName)
        {
            string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            string UploadedTaxonomyFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + TaxonomyFileName);
            string sqlTableName1, sqlTableName2 = string.Empty;
            try
            {
                #region Create temp tables and write uploaded files data
                sqlTableName1 = _BLLGAT.CreateSQLTable("spGATSTPIFCCreateInputFileTempTable");
                bool IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(InputFileName, sqlTableName1, 0);
                if (IsSucceeded) 
                {
                    sqlTableName2 = _BLLGAT.CreateSQLTable("spGATSTPIFCCreateTaxonomyFileTempTable");
                    IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(TaxonomyFileName, sqlTableName2, 0);
                }
                #endregion

                if (File.Exists(UploadedTaxonomyFilepath))
                    File.Delete(UploadedTaxonomyFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName1 + "," + sqlTableName2);
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);
                if (File.Exists(UploadedTaxonomyFilepath))
                    File.Delete(UploadedTaxonomyFilepath);
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Split PO Text and Write it to Columns
        [HttpPost]
        [Route("SplitPOTextAndWriteOutputToExcel")]
        public HttpResponseMessage SplitPOTextAndWriteOutputToExcel(string InputFileName, string UploadedInputFileName, string inputFileSqlTableName, string taxonomyFileSqlTableName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                DataTable dataTable = _BLLGATSTPInputFormatConversion.FetchInputFormatConversionDetails (inputFileSqlTableName, taxonomyFileSqlTableName);
                
                int RowCounter = 1;
                foreach(DataRow dr in dataTable.Rows)
                {
                    if (!string.IsNullOrEmpty(dr["MNFR"].ToString()))
                        wsIF.Cells[RowCounter, 5].PutValue(dr["MNFR"].ToString());
                    if (!string.IsNullOrEmpty(dr["MNFR ATTRIBUTE VALUE"].ToString()))
                        wsIF.Cells[RowCounter, 6].PutValue(dr["MNFR ATTRIBUTE VALUE"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO"].ToString()))
                        wsIF.Cells[RowCounter, 7].PutValue(dr["PTNO"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 8].PutValue(dr["PTNO Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["MNFR 1"].ToString()))
                        wsIF.Cells[RowCounter, 9].PutValue(dr["MNFR 1"].ToString());
                    if (!string.IsNullOrEmpty(dr["MNFR 1 Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 10].PutValue(dr["MNFR 1 Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO 1"].ToString()))
                        wsIF.Cells[RowCounter, 11].PutValue(dr["PTNO 1"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO 1 Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 12].PutValue(dr["PTNO 1 Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["MNFR 2"].ToString()))
                        wsIF.Cells[RowCounter, 13].PutValue(dr["MNFR 2"].ToString());
                    if (!string.IsNullOrEmpty(dr["MNFR 2 Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 14].PutValue(dr["MNFR 2 Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO 2"].ToString()))
                        wsIF.Cells[RowCounter, 15].PutValue(dr["PTNO 2"].ToString());
                    if (!string.IsNullOrEmpty(dr["PTNO 2 Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 16].PutValue(dr["PTNO 2 Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["DWG"].ToString()))
                        wsIF.Cells[RowCounter, 17].PutValue(dr["DWG"].ToString());
                    if (!string.IsNullOrEmpty(dr["DWG Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 18].PutValue(dr["DWG Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["Drawing Version"].ToString()))
                        wsIF.Cells[RowCounter, 19].PutValue(dr["Drawing Version"].ToString());
                    if (!string.IsNullOrEmpty(dr["Drawing Version Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 20].PutValue(dr["Drawing Version Attribute Value"].ToString());
                    if (!string.IsNullOrEmpty(dr["POS/TAG"].ToString()))
                        wsIF.Cells[RowCounter, 21].PutValue(dr["POS/TAG"].ToString());
                    if (!string.IsNullOrEmpty(dr["POS/TAG Attribute Value"].ToString()))
                        wsIF.Cells[RowCounter, 22].PutValue(dr["POS/TAG Attribute Value"].ToString());
                    for (int attCtr = 1; attCtr <= 36; attCtr++)
                    {
                        if (!string.IsNullOrEmpty(dr["Attribute Name " + attCtr.ToString()].ToString()))
                            wsIF.Cells[RowCounter, 22 + (attCtr * 2) - 1].PutValue(dr["Attribute Name " + attCtr.ToString()].ToString());
                        if (!string.IsNullOrEmpty(dr["Attribute Value " + attCtr.ToString()].ToString()))
                            wsIF.Cells[RowCounter, 22 + (attCtr * 2)].PutValue(dr["Attribute Value " + attCtr.ToString()].ToString());
                    }
                    if (!string.IsNullOrEmpty(dr["Remaining_Information"].ToString()))
                        wsIF.Cells[RowCounter, 95].PutValue(dr["Remaining_Information"].ToString());

                    RowCounter++;
                }

                wsIF.AutoFitColumns();
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
