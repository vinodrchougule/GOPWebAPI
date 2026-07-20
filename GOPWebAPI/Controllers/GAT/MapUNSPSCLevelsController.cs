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
using System.Xml.Linq;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/MapUNSPSCLevels")]
    public class MapUNSPSCLevelsController : ApiController
    {
        private BLLGAT _BLLGAT;
        private BLLMapUNSPSCLevels _BLLMapUNSPSCLevels;
        public MapUNSPSCLevelsController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGAT = new BLLGAT(connectionString);
            _BLLMapUNSPSCLevels = new BLLMapUNSPSCLevels(connectionString);
        }

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

                #region Check uploaded file first worksheet has UNSPSC Code column
                int UNSPSCCodeColNo = -1;
                for(int col=0;col<=wsIFFW.Cells.MaxColumn;col++)
                {
                    if (wsIFFW.Cells[0,col].StringValue.Trim().ToLower()=="unspsc code")
                    {
                        UNSPSCCodeColNo = col; break;
                    }
                }

                if (UNSPSCCodeColNo<0)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no 'UNSPSC Code' column.");
                }

                if (wsIFFW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                List<GATMapUNSPSC> UNSPSCDataList = new List<GATMapUNSPSC>();
                for (int row = 1; row <= wsIFFW.Cells.MaxRow; row++)
                {
                    UNSPSCDataList.Add(new GATMapUNSPSC
                    {
                        SLNO = row.ToString(),
                        UNSPSCCode = wsIFFW.Cells[row, UNSPSCCodeColNo].StringValue.Trim()
                    });
                }

                #region Create temp table and write uploaded file data
                string sqlTableName = _BLLGAT.CreateSQLTable("spMapUNSPSCLevelCreateTempTable");
                bool IsSucceeded = false;

                if (!string.IsNullOrEmpty(sqlTableName)) 
                    IsSucceeded = _BLLGAT.WriteListDataToSQLServerTable(UNSPSCDataList, sqlTableName);
                #endregion

                if (IsSucceeded)
                    return Request.CreateResponse(HttpStatusCode.OK, sqlTableName);
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Failed to write input file unspsc code data to database");
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

        #region Map UNSPSC Levels and write output to excel
        [HttpPost]
        [Route("MapUNSPSCLevelsAndWriteOutputToExcel")]
        public HttpResponseMessage MapUNSPSCLevelsAndWriteOutputToExcel(string UploadedInputFileName,string InputFileName, string sqlTableName, string UNSPSCVersion)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                DataTable dataTable = _BLLMapUNSPSCLevels.MapUNSPSCLevels(sqlTableName, UNSPSCVersion);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);
                Worksheet wsIFFW = wbIF.Worksheets[0];      //Input File First Worksheet

                int UNSPSCCodeColNo = -1;
                for (int col = 0; col <= wsIFFW.Cells.MaxColumn; col++)
                {
                    if (wsIFFW.Cells[0, col].StringValue.Trim().ToLower() == "unspsc code")
                    {
                        UNSPSCCodeColNo = col; break;
                    }
                }

                int SegmentColNo = UNSPSCCodeColNo, FamilyColNo = SegmentColNo + 1, ClassColNo = FamilyColNo + 1, CommodityColNo = ClassColNo + 1;
                wsIFFW.Cells.InsertColumn(UNSPSCCodeColNo);
                wsIFFW.Cells.InsertColumn(UNSPSCCodeColNo);
                wsIFFW.Cells.InsertColumn(UNSPSCCodeColNo);
                wsIFFW.Cells.InsertColumn(UNSPSCCodeColNo);

                wsIFFW.Cells[0, SegmentColNo].PutValue("Segment");
                wsIFFW.Cells[0, FamilyColNo].PutValue("Family");
                wsIFFW.Cells[0, ClassColNo].PutValue("Class");
                wsIFFW.Cells[0, CommodityColNo].PutValue("Commodity");

                #region Setting Styles
                AsposeHelpers asposeHelpers = new AsposeHelpers();
                Aspose.Cells.Style styleLeftAlignData = asposeHelpers.GetStyle(wbIF, 0, "left");
                #endregion

                #region Writing row data
                int RowCounter = 1;
                foreach (DataRow dr in dataTable.Rows)
                {
                    wsIFFW.Cells[RowCounter, SegmentColNo].PutValue(dr["Segment"]);
                    wsIFFW.Cells[RowCounter, FamilyColNo].PutValue(dr["Family"]);
                    wsIFFW.Cells[RowCounter, ClassColNo].PutValue(dr["Class"]);
                    wsIFFW.Cells[RowCounter, CommodityColNo].PutValue(dr["Commodity"]);

                    #region setting row data style
                    wsIFFW.Cells[RowCounter, SegmentColNo].SetStyle(styleLeftAlignData);
                    wsIFFW.Cells[RowCounter, FamilyColNo].SetStyle(styleLeftAlignData);
                    wsIFFW.Cells[RowCounter, ClassColNo].SetStyle(styleLeftAlignData);
                    wsIFFW.Cells[RowCounter, CommodityColNo].SetStyle(styleLeftAlignData);
                    #endregion

                    RowCounter++;
                }
                #endregion

                wsIFFW.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);

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
    }
}
