using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Xml;
using System.Xml.Serialization;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/PeriodicProjectReport")]
    public class PeriodicProjectReportController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public PeriodicProjectReportController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Periodic Project Report data
        [HttpPost]
        [Route("ReadPeriodicProjectReportData")]
        public IHttpActionResult ReadPeriodicProjectReportData([FromBody] PeriodicProjectReportInputModel periodicProjectReportInputModel)
        {
            try
            {
                int SlNo = 1;

                //Swagger fix
                if (periodicProjectReportInputModel.BatchNo == "{BatchNo}")
                    periodicProjectReportInputModel.BatchNo = "";

                //Create a list to hold the Project Report Data
                List<PeriodicProjectReportModel> ProjectReportList = new List<PeriodicProjectReportModel>();

                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spPeriodicProjectReport";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", periodicProjectReportInputModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", periodicProjectReportInputModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", periodicProjectReportInputModel.BatchNo);
                    cmd.Parameters.AddWithValue("@FromDate", periodicProjectReportInputModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", periodicProjectReportInputModel.ToDate);

                    //Calling sp to the report data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        PeriodicProjectReportModel periodicProjectReportModel = new PeriodicProjectReportModel();

                        periodicProjectReportModel.SlNo = SlNo;
                        periodicProjectReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                        periodicProjectReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                        periodicProjectReportModel.Activity = sqlReader["Activity"].ToString();
                        periodicProjectReportModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                        periodicProjectReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        periodicProjectReportModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                        periodicProjectReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);

                        ProjectReportList.Add(periodicProjectReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectReportList);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Export Periodic Project Report data to Excel
        [HttpPost]
        [Route("ExportPeriodicProjectReportDataToExcel")]
        public HttpResponseMessage ExportPeriodicProjectReportDataToExcel([FromBody] PeriodicProjectReportInputModel periodicProjectReportInputModel)
        {
            try
            {
                if (periodicProjectReportInputModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Periodic Project Report", periodicProjectReportInputModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = string.Empty;

                    //Swagger fix
                    if (periodicProjectReportInputModel.BatchNo == "{BatchNo}")
                        periodicProjectReportInputModel.BatchNo = "";

                    if (string.IsNullOrEmpty(periodicProjectReportInputModel.BatchNo))
                        FileName = "PeriodicProjectReport_" + periodicProjectReportInputModel.CustomerCode + '_' + periodicProjectReportInputModel.ProjectCode + ".xlsx";
                    else
                        FileName = "PeriodicProjectReport_" + periodicProjectReportInputModel.CustomerCode + '_' + periodicProjectReportInputModel.ProjectCode + '_' + periodicProjectReportInputModel.BatchNo + ".xlsx";

                    //Create a list to hold the Project Report Data
                    List<PeriodicProjectReportModel> ProjectReportList = new List<PeriodicProjectReportModel>();

                    System.Data.Common.DbDataReader sqlReader;
                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    var ws = wb.Worksheets[0];
                    int row = 8, counter = 1;

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignDataWithoutBorder = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleRightAlignData = ws.Cells[0, 0].GetStyle();

                    styleHeader.IsTextWrapped = true;
                    styleHeader.HorizontalAlignment = TextAlignmentType.Center;
                    styleHeader.VerticalAlignment = TextAlignmentType.Center;
                    styleHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                    styleHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                    styleHeader.Pattern = BackgroundType.VerticalStripe;
                    styleHeader.Font.Color = System.Drawing.Color.Black;
                    styleHeader.Font.IsBold = true;
                    styleHeader.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    styleCenterAlignData.HorizontalAlignment = TextAlignmentType.Center;
                    styleCenterAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    styleLeftAlignDataWithoutBorder.HorizontalAlignment = TextAlignmentType.Left;

                    styleLeftAlignData.HorizontalAlignment = TextAlignmentType.Left;
                    styleLeftAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    #endregion

                    using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                    {
                        //Initialize command object
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "spPeriodicProjectReport";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", periodicProjectReportInputModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", periodicProjectReportInputModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", periodicProjectReportInputModel.BatchNo);
                        cmd.Parameters.AddWithValue("@FromDate", periodicProjectReportInputModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", periodicProjectReportInputModel.ToDate);

                        //Calling sp to the report data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                PeriodicProjectReportModel periodicProjectReportModel = new PeriodicProjectReportModel();

                                periodicProjectReportModel.SlNo = SlNo;
                                periodicProjectReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                                periodicProjectReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                                periodicProjectReportModel.Activity = sqlReader["Activity"].ToString();
                                periodicProjectReportModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                                periodicProjectReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                periodicProjectReportModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                                periodicProjectReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);

                                ProjectReportList.Add(periodicProjectReportModel);
                                SlNo++;
                            }
                            #endregion
                            conn.Close();

                            #region Writing report header
                            for (int col = 0; col <= 7; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 8);
                            ws.Cells[1, 0].PutValue("Periodic Project Report");

                            for (int col = 0; col <= 1; col++)
                                ws.Cells[3, col].SetStyle(styleLeftAlignDataWithoutBorder);
                            ws.Cells.Merge(3, 0, 1, 2);
                            ws.Cells[3, 0].PutValue("Customer Code: " + periodicProjectReportInputModel.CustomerCode);

                            for (int col = 3; col <= 4; col++)
                                ws.Cells[3, col].SetStyle(styleLeftAlignDataWithoutBorder);
                            ws.Cells.Merge(3, 3, 1, 2);
                            ws.Cells[3, 3].PutValue("Project Code: " + periodicProjectReportInputModel.ProjectCode);

                            if (!string.IsNullOrEmpty(periodicProjectReportInputModel.BatchNo))
                            {
                                for (int col = 6; col <= 7; col++)
                                    ws.Cells[3, col].SetStyle(styleLeftAlignDataWithoutBorder);
                                ws.Cells.Merge(3, 6, 1, 2);
                                ws.Cells[3, 6].PutValue("Batch No.: " + periodicProjectReportInputModel.BatchNo);
                            }

                            for (int col = 0; col <= 1; col++)
                                ws.Cells[5, col].SetStyle(styleLeftAlignDataWithoutBorder);
                            ws.Cells.Merge(5, 0, 1, 2);
                            ws.Cells[5, 0].PutValue("From Date: " + periodicProjectReportInputModel.FromDate.ToString("dd-MMM-yyyy"));

                            for (int col = 3; col <= 4; col++)
                                ws.Cells[5, col].SetStyle(styleLeftAlignDataWithoutBorder);
                            ws.Cells.Merge(5, 3, 1, 2);
                            ws.Cells[5, 3].PutValue("To Date: " + periodicProjectReportInputModel.ToDate.ToString("dd-MMM-yyyy"));
                            #endregion

                            #region Writing column headings
                            ws.Cells[7, 0].PutValue("S.No.");
                            ws.Cells[7, 1].PutValue("Employee Code");
                            ws.Cells[7, 2].PutValue("Employee Name");
                            ws.Cells[7, 3].PutValue("Activity");
                            ws.Cells[7, 4].PutValue("Production Allocated Count");
                            ws.Cells[7, 5].PutValue("Production Completed Count");
                            ws.Cells[7, 6].PutValue("QC Allocated Count");
                            ws.Cells[7, 7].PutValue("QC Completed Count");

                            for (int c = 0; c <= 7; c++)
                                ws.Cells[7, c].SetStyle(styleHeader);
                            #endregion

                            foreach (PeriodicProjectReportModel periodicProjectReportModel in ProjectReportList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(counter);
                                ws.Cells[row, 1].PutValue(periodicProjectReportModel.EmployeeCode);
                                ws.Cells[row, 2].PutValue(periodicProjectReportModel.EmployeeName);
                                ws.Cells[row, 3].PutValue(periodicProjectReportModel.Activity);
                                ws.Cells[row, 4].PutValue(periodicProjectReportModel.ProductionAllocatedCount);
                                ws.Cells[row, 5].PutValue(periodicProjectReportModel.ProductionCompletedCount);
                                ws.Cells[row, 6].PutValue(periodicProjectReportModel.QCAllocatedCount);
                                ws.Cells[row, 7].PutValue(periodicProjectReportModel.QCCompletedCount);
                                #endregion

                                #region setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                                #endregion

                                row++;
                                counter++;
                            }

                            #region Saving and downloading the report
                            ws.AutoFitColumns();
                            string tempPath = System.IO.Path.GetTempPath();
                            string filename = tempPath + FileName;
                            wb.Save(filename);

                            //Read the file into a Byte Array.
                            byte[] bytes = File.ReadAllBytes(filename);

                            //Set the Response Content.
                            response.Content = new ByteArrayContent(bytes);

                            //Set the Response Content Length.
                            response.Content.Headers.ContentLength = bytes.LongLength;

                            //Set the Content Disposition Header Value and FileName.
                            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                            response.Content.Headers.ContentDisposition.FileName = FileName;

                            //Set the File Content Type.
                            response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(filename));

                            return response;
                            #endregion
                        }
                        else
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No data found");
                    }
                }
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
