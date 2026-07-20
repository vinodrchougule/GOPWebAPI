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
    [RoutePrefix("api/SamplingSelection")]
    public class SamplingSelectionController : ApiController
    {
        private BLLGAT _BLLGAT;
        private BLLGATSamplingSelection _BLLGATSamplingSelection;

        public SamplingSelectionController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
            _BLLGATSamplingSelection = new BLLGATSamplingSelection(connectionString);
        }

        #region Validate the input file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName, int SamplingPercentage)
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

                if(SamplingPercentage > 99 || SamplingPercentage <=0)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid sampling percentage");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Sl No' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "MATERIAL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Material No' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells.MaxRow <= 0)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file first Worksheet has no data rows.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully");
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

        #region Check Duplicate Material No. exists in input file
        [HttpPost]
        [Route("CheckDuplicateMaterialNoExists")]
        public HttpResponseMessage CheckDuplicateMaterialNoExists(string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                #region Create temp table and write uploaded file data
                bool IsSucceeded = false;
                string sqlTableName = _BLLGAT.CreateSQLTable("spGATSamplingSelectionTempTable");
                if(!string.IsNullOrEmpty(sqlTableName))
                    IsSucceeded = _BLLGAT.WriteFileDataToSQLServerTable(InputFileName, sqlTableName, 0);
                #endregion

                if(IsSucceeded)
                {
                    string CheckDuplicateMaterialNoExists = _BLLGATSamplingSelection.CheckDuplicateMaterialNoExists(sqlTableName);
                    if(CheckDuplicateMaterialNoExists.Trim().ToLower()!="success")
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, CheckDuplicateMaterialNoExists);
                }

                return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
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

        #region Fetching the data from Database and Write to Excel
        [HttpPost]
        [Route("SelectRandomRowsAndWriteToOutputFile")]
        public HttpResponseMessage SelectRandomRowsAndWriteToOutputFile(string UploadedInputFileName, string InputFileName,string sqlTableName, int SamplingSelectionPercentage)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                DataTable dataTable = _BLLGATSamplingSelection.SelectAndFetchRandomRows(sqlTableName, SamplingSelectionPercentage);

                #region Setting up the workbook
                AsposeHelpers asposeHelpers = new AsposeHelpers();
                Workbook wb = asposeHelpers.GetWorkbook();
                var ws1 = wb.Worksheets[0];
                int row = 1;
                #endregion

                #region Setting styles
                Aspose.Cells.Style styleHeader = ws1.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleCenterAlignData = asposeHelpers.GetStyle(wb, 0, "center");

                styleHeader.IsTextWrapped = true;
                styleHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleHeader.VerticalAlignment = TextAlignmentType.Center;
                styleHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                styleHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                styleHeader.Pattern = BackgroundType.VerticalStripe;
                styleHeader.Font.Color = System.Drawing.Color.Black;
                styleHeader.Font.IsBold = true;
                #endregion

                #region Write Random selected rows to first worksheet
                #region Write header to first worksheet
                ws1.Cells[0, 0].PutValue("SLNO");
                ws1.Cells[0, 1].PutValue("MaterialNo");
                ws1.Cells[0, 2].PutValue("Noun");

                for (int c = 0; c <= 2; c++)
                    ws1.Cells[0, c].SetStyle(styleHeader);
                #endregion

                #region Writing row data
                foreach (DataRow dr in dataTable.Rows)
                {
                    #region Writing row data
                    ws1.Cells[row, 0].PutValue(dr["SLNO"].ToString().Trim());
                    ws1.Cells[row, 1].PutValue(dr["MaterialNo"].ToString().Trim());
                    ws1.Cells[row, 2].PutValue(dr["Noun"].ToString().Trim());
                    #endregion

                    #region setting row data style
                    ws1.Cells[row, 0].SetStyle(styleCenterAlignData);
                    ws1.Cells[row, 1].SetStyle(styleCenterAlignData);
                    ws1.Cells[row, 2].SetStyle(styleCenterAlignData);
                    #endregion

                    row++;
                }
                #endregion
                #endregion

                #region Write Items Count for each Noun from input file data in second worksheet
                DataTable dtNounItemsCount = _BLLGATSamplingSelection.FetchItemsCountForEachNoun(sqlTableName);
                Worksheet ws2 = wb.Worksheets[1];

                #region Write headers to second worksheet
                ws2.Cells[0, 0].PutValue("Noun");
                ws2.Cells[0, 1].PutValue("Count");
                #endregion

                for (int c = 0; c <= 1; c++)
                    ws2.Cells[0, c].SetStyle(styleHeader);

                #region Write row data to second worksheet
                row = 1;
                foreach (DataRow dr in dtNounItemsCount.Rows)
                {
                    #region Writing row data
                    ws2.Cells[row, 0].PutValue(dr["Noun"].ToString().Trim());
                    ws2.Cells[row, 1].PutValue(dr["CountofItems"].ToString().Trim());
                    #endregion

                    ws2.Cells[row, 0].SetStyle(styleCenterAlignData);
                    ws2.Cells[row, 1].SetStyle(styleCenterAlignData);

                    row++;
                }
                #endregion
                #endregion

                #region Write the Count of Items for each Noun from randomly selected items to third worksheet
                DataTable dtRandomItemsCountForEachNoun = _BLLGATSamplingSelection.FetchItemsCountForEachNounFromRandomlySelectedItems(sqlTableName, SamplingSelectionPercentage);
                var ws3 = wb.Worksheets[2];

                #region Write headers to third worksheet
                ws3.Cells[0, 0].PutValue("Noun");
                ws3.Cells[0, 1].PutValue("Count");
                #endregion

                for (int c = 0; c <= 1; c++)
                    ws3.Cells[0, c].SetStyle(styleHeader);

                #region Write row data to third worksheet
                row = 1;
                foreach (DataRow dr in dtRandomItemsCountForEachNoun.Rows)
                {
                    #region Writing row data
                    ws3.Cells[row, 0].PutValue(dr["Noun"].ToString().Trim());
                    ws3.Cells[row, 1].PutValue(dr["CountofItems"].ToString().Trim());
                    #endregion

                    ws3.Cells[row, 0].SetStyle(styleCenterAlignData);
                    ws3.Cells[row, 1].SetStyle(styleCenterAlignData);

                    row++;
                }
                #endregion
                #endregion

                #region Save the file
                ws1.AutoFitColumns();ws2.AutoFitColumns();ws3.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wb.Save(filename);
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
    }
}