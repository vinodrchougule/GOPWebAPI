using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using System.Data.SqlTypes;
using System.Net.Http.Headers;
using GOPWebAPI.Helpers;
using Aspose.Cells;
using System.Data.OleDb;
using System.Data.Common;
using GOPWebAPI.BLL;
using System.Configuration;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/EmployeeSpecificReport")]
    public class EmployeeSpecificReportController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public EmployeeSpecificReportController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Employee Specific Details Report data
        [HttpGet]
        [Route("ReadEmployeeSpecificDetailsReportData/{EmployeeCode}/{FromDate}/{ToDate}")]
        public IHttpActionResult ReadEmployeeSpecificDetailsReportData(string EmployeeCode, DateTime FromDate, DateTime ToDate)
        {
            try
            {
                int SlNo = 1;

                if (EmployeeCode.Contains('-'))
                {
                    string[] arrEmployeeNameCode = EmployeeCode.Split('-');
                    EmployeeCode = arrEmployeeNameCode[1].Trim();
                }

                //Create a list to hold the Employee Specific Report Data
                List<EmployeeSpecificDetailsReportModel> EmployeeSpecificDetailsReportList = new List<EmployeeSpecificDetailsReportModel>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spEmployeeSpecificReportDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@EmployeeCode", EmployeeCode);
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get list of Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        EmployeeSpecificDetailsReportModel employeeSpecificDetailsReportModel = new EmployeeSpecificDetailsReportModel();

                        employeeSpecificDetailsReportModel.SlNo = SlNo;
                        employeeSpecificDetailsReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                        employeeSpecificDetailsReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                        employeeSpecificDetailsReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                        employeeSpecificDetailsReportModel.Activity = sqlReader["Activity"].ToString();
                        employeeSpecificDetailsReportModel.ProductionOrQCCompletedOn = Convert.ToDateTime(sqlReader["ProductionOrQCCompletedOn"]);
                        employeeSpecificDetailsReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        employeeSpecificDetailsReportModel.ProductionTarget = Convert.ToInt32(sqlReader["ProductionTarget"]);
                        employeeSpecificDetailsReportModel.ProductionProductiveHours = Convert.ToDecimal(sqlReader["ProductionProductiveHours"]);
                        employeeSpecificDetailsReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        employeeSpecificDetailsReportModel.QCTarget = Convert.ToInt32(sqlReader["QCTarget"]);
                        employeeSpecificDetailsReportModel.QCProductiveHours = Convert.ToDecimal(sqlReader["QCProductiveHours"]);

                        EmployeeSpecificDetailsReportList.Add(employeeSpecificDetailsReportModel);
                        SlNo++;
                    }
                    conn.Close();
                    
                    //return list to the request
                    return Ok(EmployeeSpecificDetailsReportList);
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

        #region Read Employee Specific Summary Report data
        [HttpGet]
        [Route("ReadEmployeeSpecificSummaryReportData/{EmployeeCode}/{FromDate}/{ToDate}")]
        public IHttpActionResult ReadEmployeeSpecificSummaryReportData(string EmployeeCode, DateTime FromDate, DateTime ToDate)
        {
            try
            {
                int SlNo = 1;

                if (EmployeeCode.Contains('-'))
                {
                    string[] arrEmployeeNameCode = EmployeeCode.Split('-');
                    EmployeeCode = arrEmployeeNameCode[1].Trim();
                }

                //Create a list to hold the Employee Specific Report Data
                List<EmployeeSpecificSummaryReportModel> employeeSpecificSummaryReportList = new List<EmployeeSpecificSummaryReportModel>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spEmployeeSpecificReportSummary";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@EmployeeCode", EmployeeCode);
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get list of Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        EmployeeSpecificSummaryReportModel employeeSpecificSummaryReportModel = new EmployeeSpecificSummaryReportModel();

                        employeeSpecificSummaryReportModel.SlNo = SlNo;
                        employeeSpecificSummaryReportModel.ProductionOrQCCompletedOn = Convert.ToDateTime(sqlReader["ProductionOrQCCompletedOn"]);
                        employeeSpecificSummaryReportModel.ProductiveHours = Convert.ToDecimal(sqlReader["ProductiveHours"]);

                        employeeSpecificSummaryReportList.Add(employeeSpecificSummaryReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(employeeSpecificSummaryReportList);
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

        #region Export Employee Specific Details Report data to Excel
        [HttpGet]
        [Route("ExportEmployeeSpecificDetailsReportDataToExcel/{EmployeeCode}/{FromDate}/{ToDate}/{UserID}")]
        public HttpResponseMessage ExportEmployeeSpecificDetailsReportDataToExcel(string EmployeeCode, DateTime FromDate, DateTime ToDate, string UserID)
        {
             try
             {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Employee Specific Details Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "Employee Specific Details Report.xlsx";
                    string EmployeeName = string.Empty;

                    if (EmployeeCode.Contains('-'))
                    {
                        string[] arrEmployeeNameCode = EmployeeCode.Split('-');
                        EmployeeName = arrEmployeeNameCode[0].Trim();
                        EmployeeCode = arrEmployeeNameCode[1].Trim();
                    }

                    //Create a list to hold the Employee Specific Report Data
                    List<EmployeeSpecificDetailsReportModel> employeeSpecificDetailsReportList = new List<EmployeeSpecificDetailsReportModel>();

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
                    int row = 6;
                    #endregion

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();

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
                        cmd.CommandText = "spEmployeeSpecificReportDetails";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@EmployeeCode", EmployeeCode);
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);

                        //Calling sp to get list of Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                EmployeeSpecificDetailsReportModel employeeSpecificDetailsReportModel = new EmployeeSpecificDetailsReportModel();

                                employeeSpecificDetailsReportModel.SlNo = SlNo;
                                employeeSpecificDetailsReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                                employeeSpecificDetailsReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                                employeeSpecificDetailsReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                                employeeSpecificDetailsReportModel.Activity = sqlReader["Activity"].ToString();
                                employeeSpecificDetailsReportModel.ProductionOrQCCompletedOn = Convert.ToDateTime(sqlReader["ProductionOrQCCompletedOn"]);
                                employeeSpecificDetailsReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                employeeSpecificDetailsReportModel.ProductionTarget = Convert.ToInt32(sqlReader["ProductionTarget"]);
                                employeeSpecificDetailsReportModel.ProductionProductiveHours = Convert.ToDecimal(sqlReader["ProductionProductiveHours"]);
                                employeeSpecificDetailsReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                                employeeSpecificDetailsReportModel.QCTarget = Convert.ToInt32(sqlReader["QCTarget"]);
                                employeeSpecificDetailsReportModel.QCProductiveHours = Convert.ToDecimal(sqlReader["QCProductiveHours"]);

                                employeeSpecificDetailsReportList.Add(employeeSpecificDetailsReportModel);
                                SlNo++;
                            }
                            #endregion
                            conn.Close();

                            #region Writing report header
                            for (int col = 0; col <= 11; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 12);
                            ws.Cells[1, 0].PutValue("Employee Specific Detail Projects Report");

                            ws.Cells.Merge(3, 0, 1, 3);
                            ws.Cells[3, 0].PutValue("Employee Name: " + EmployeeName);
                            ws.Cells[3, 4].PutValue("From Date: " + FromDate.ToString("dd-MMM-yyyy"));
                            ws.Cells[3, 6].PutValue("To Date: " + ToDate.ToString("dd-MMM-yyyy"));
                            #endregion

                            #region Writing column headings
                            ws.Cells[5, 0].PutValue("S.No.");
                            ws.Cells[5, 1].PutValue("Customer Code");
                            ws.Cells[5, 2].PutValue("Project Code");
                            ws.Cells[5, 3].PutValue("Batch No.");
                            ws.Cells[5, 4].PutValue("Activity");
                            ws.Cells[5, 5].PutValue("Production or QC Completed on");
                            ws.Cells[5, 6].PutValue("Production Completed Count");
                            ws.Cells[5, 7].PutValue("Production Target");
                            ws.Cells[5, 8].PutValue("Production Productive Hours");
                            ws.Cells[5, 9].PutValue("QC Completed Count");
                            ws.Cells[5, 10].PutValue("QC Target");
                            ws.Cells[5, 11].PutValue("QC Productive Hours");

                            for (int c = 0; c <= 11; c++)
                                ws.Cells[5, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (EmployeeSpecificDetailsReportModel employeeSpecificDetailsReportModel in employeeSpecificDetailsReportList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(employeeSpecificDetailsReportModel.SlNo);
                                ws.Cells[row, 1].PutValue(employeeSpecificDetailsReportModel.CustomerCode);
                                ws.Cells[row, 2].PutValue(employeeSpecificDetailsReportModel.ProjectCode);
                                if (!string.IsNullOrEmpty(employeeSpecificDetailsReportModel.BatchNo))
                                    ws.Cells[row, 3].PutValue(employeeSpecificDetailsReportModel.BatchNo);
                                ws.Cells[row, 4].PutValue(employeeSpecificDetailsReportModel.Activity);
                                ws.Cells[row, 5].PutValue(employeeSpecificDetailsReportModel.ProductionOrQCCompletedOn.ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 6].PutValue(employeeSpecificDetailsReportModel.ProductionCompletedCount);
                                ws.Cells[row, 7].PutValue(employeeSpecificDetailsReportModel.ProductionTarget);
                                ws.Cells[row, 8].PutValue(employeeSpecificDetailsReportModel.ProductionProductiveHours);
                                ws.Cells[row, 9].PutValue(employeeSpecificDetailsReportModel.QCCompletedCount);
                                ws.Cells[row, 10].PutValue(employeeSpecificDetailsReportModel.QCTarget);
                                ws.Cells[row, 11].PutValue(employeeSpecificDetailsReportModel.QCProductiveHours);
                                #endregion

                                #region Setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 11].SetStyle(styleCenterAlignData);
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

        #region Export Employee Specific Summary Report data to Excel
        [HttpGet]
        [Route("ExportEmployeeSpecificSummaryReportDataToExcel/{EmployeeCode}/{FromDate}/{ToDate}/{UserID}")]
        public HttpResponseMessage ExportEmployeeSpecificSummaryReportDataToExcel(string EmployeeCode, DateTime FromDate, DateTime ToDate, string UserID)
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Employee Specific Summary Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "Employee Specific Summary Report.xlsx";
                    string EmployeeName = string.Empty;

                    if (EmployeeCode.Contains('-'))
                    {
                        string[] arrEmployeeNameCode = EmployeeCode.Split('-');
                        EmployeeName = arrEmployeeNameCode[0].Trim();
                        EmployeeCode = arrEmployeeNameCode[1].Trim();
                    }

                    //Create a list to hold the Employee Specific Report Data
                    List<EmployeeSpecificSummaryReportModel> employeeSpecificSummaryReportList = new List<EmployeeSpecificSummaryReportModel>();

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
                    int row = 6;
                    #endregion

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();

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
                        cmd.CommandText = "spEmployeeSpecificReportSummary";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@EmployeeCode", EmployeeCode);
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);

                        //Calling sp to get list of Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                EmployeeSpecificSummaryReportModel employeeSpecificSummaryReportModel = new EmployeeSpecificSummaryReportModel();

                                employeeSpecificSummaryReportModel.SlNo = SlNo;
                                employeeSpecificSummaryReportModel.ProductionOrQCCompletedOn = Convert.ToDateTime(sqlReader["ProductionOrQCCompletedOn"]);
                                employeeSpecificSummaryReportModel.ProductiveHours = Convert.ToDecimal(sqlReader["ProductiveHours"]);

                                employeeSpecificSummaryReportList.Add(employeeSpecificSummaryReportModel);
                                SlNo++;
                            }
                            conn.Close();
                            #endregion

                            #region Writing report header
                            for (int col = 0; col <= 6; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 7);
                            ws.Cells[1, 0].PutValue("Employee Specific Summary Projects Report");

                            ws.Cells.Merge(3, 0, 1, 3);
                            ws.Cells[3, 0].PutValue("Employee Name: " + EmployeeName);
                            ws.Cells[3, 4].PutValue("From Date: " + FromDate.ToString("dd-MMM-yyyy"));
                            ws.Cells[3, 6].PutValue("To Date: " + ToDate.ToString("dd-MMM-yyyy"));
                            #endregion

                            #region Writing column headings
                            ws.Cells[5, 1].PutValue("S.No.");
                            ws.Cells[5, 2].PutValue("Production or QC Completed on");
                            ws.Cells[5, 3].PutValue("Productive Hours");

                            for (int c = 1; c <= 3; c++)
                                ws.Cells[5, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (EmployeeSpecificSummaryReportModel employeeSpecificSummaryReportModel in employeeSpecificSummaryReportList)
                            {
                                #region Writing row data
                                ws.Cells[row, 1].PutValue(employeeSpecificSummaryReportModel.SlNo);
                                ws.Cells[row, 2].PutValue(employeeSpecificSummaryReportModel.ProductionOrQCCompletedOn.ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 3].PutValue(employeeSpecificSummaryReportModel.ProductiveHours);
                                #endregion

                                #region Setting row data style
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
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
