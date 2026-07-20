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
    [RoutePrefix("api/PreviousDayReport")]
    public class PreviousDayReportController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public PreviousDayReportController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Previous Day Report data
        [HttpGet]
        [Route("ReadPreviousDayReportData/{FromDate}/{ToDate}")]
        public IHttpActionResult ReadPreviousDayReportData(DateTime FromDate, DateTime ToDate)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the Project Previous Day Report Data
                List<PreviousDayReportModel> PreviousDayReportList = new List<PreviousDayReportModel>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectsReportWithinDateRange";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get list of Project Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        PreviousDayReportModel previousDayReportModel = new PreviousDayReportModel();

                        previousDayReportModel.SlNo = SlNo;
                        previousDayReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                        previousDayReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                        previousDayReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                        previousDayReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                        previousDayReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                        previousDayReportModel.Department = sqlReader["Department"].ToString();
                        previousDayReportModel.ManagerName = sqlReader["ManagerName"].ToString();
                        previousDayReportModel.Activity = sqlReader["Activity"].ToString();
                        previousDayReportModel.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                        previousDayReportModel.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);

                        PreviousDayReportList.Add(previousDayReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(PreviousDayReportList);
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

        #region Export Previous Day Report data to Excel
        [HttpGet]
        [Route("ExportPreviousDayReportDataToExcel/{FromDate}/{ToDate}/{UserID}")]
        public HttpResponseMessage ExportPreviousDayReportDataToExcel(DateTime FromDate, DateTime ToDate, string UserID)
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Previous Day Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "ProjectsReport.xlsx";

                    //Create a list to hold the Project Previous Day Report Data
                    List<PreviousDayReportModel> PreviousDayReportList = new List<PreviousDayReportModel>();

                    System.Data.Common.DbDataReader sqlReader;
                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Setting up the workbook
                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    var ws = wb.Worksheets[0];
                    int row = 4;
                    #endregion

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();
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
                        cmd.CommandText = "spProjectsReportWithinDateRange";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);

                        //Calling sp to get list of Project Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                PreviousDayReportModel previousDayReportModel = new PreviousDayReportModel();

                                previousDayReportModel.SlNo = SlNo;
                                previousDayReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                                previousDayReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                                previousDayReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                                previousDayReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                                previousDayReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                                previousDayReportModel.Department = sqlReader["Department"].ToString();
                                previousDayReportModel.ManagerName = sqlReader["ManagerName"].ToString();
                                previousDayReportModel.Activity = sqlReader["Activity"].ToString();
                                previousDayReportModel.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                                previousDayReportModel.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);

                                PreviousDayReportList.Add(previousDayReportModel);
                                SlNo++;
                            }
                            #endregion
                            conn.Close();

                            #region Writing report header
                            for (int col = 0; col <= 10; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 11);
                            ws.Cells[1, 0].PutValue("Projects Report From " + FromDate.ToString("dd-MMM-yyyy") + " To " + ToDate.ToString("dd-MMM-yyyy"));

                            #endregion

                            #region Writing column headings
                            ws.Cells[3, 0].PutValue("S.No.");
                            ws.Cells[3, 1].PutValue("Customer Code");
                            ws.Cells[3, 2].PutValue("Project Code");
                            ws.Cells[3, 3].PutValue("Batch No.");
                            ws.Cells[3, 4].PutValue("Employee Code");
                            ws.Cells[3, 5].PutValue("Employee Name");
                            ws.Cells[3, 6].PutValue("Department");
                            ws.Cells[3, 7].PutValue("Manager");
                            ws.Cells[3, 8].PutValue("Activity");
                            ws.Cells[3, 9].PutValue("Production Completed Count");
                            ws.Cells[3, 10].PutValue("QC Completed Count");

                            for (int c = 0; c <= 10; c++)
                                ws.Cells[3, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (PreviousDayReportModel previousDayReportModel in PreviousDayReportList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(previousDayReportModel.SlNo);
                                ws.Cells[row, 1].PutValue(previousDayReportModel.CustomerCode);
                                ws.Cells[row, 2].PutValue(previousDayReportModel.ProjectCode);
                                ws.Cells[row, 3].PutValue(previousDayReportModel.BatchNo);
                                ws.Cells[row, 4].PutValue(previousDayReportModel.EmployeeCode);
                                ws.Cells[row, 5].PutValue(previousDayReportModel.EmployeeName);
                                ws.Cells[row, 6].PutValue(previousDayReportModel.Department);
                                ws.Cells[row, 7].PutValue(previousDayReportModel.ManagerName);
                                ws.Cells[row, 8].PutValue(previousDayReportModel.Activity);
                                ws.Cells[row, 9].PutValue(previousDayReportModel.ProductionCompletedCount);
                                ws.Cells[row, 10].PutValue(previousDayReportModel.QCCompletedCount);
                                #endregion

                                #region setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 8].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                #endregion                            
                                row++;
                            }
                            #endregion

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
