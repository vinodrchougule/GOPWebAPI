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
    [RoutePrefix("api/EmployeesTaskReport")]
    public class EmployeesTaskReportController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public EmployeesTaskReportController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Read Employees Task Details Report data
        [HttpPost]
        [Route("ReadEmployeesTaskDetailsReportData")]
        public IHttpActionResult ReadEmployeesTaskDetailsReportData([FromBody] EmployeesTaskReportModel employeesTaskReportModel)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the Employees Task Details Report data
                List<EmployeesTaskReportModel> employeesTaskReportDataList = new List<EmployeesTaskReportModel>();
                
                #region Employee Codes
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<EmployeeCodes>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, employeesTaskReportModel.EmployeeCodes);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spEmployeesTaskReport";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", employeesTaskReportModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", employeesTaskReportModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", employeesTaskReportModel.BatchNo);
                    cmd.Parameters.AddWithValue("@Department", employeesTaskReportModel.Department);
                    //cmd.Parameters.AddWithValue("@EmployeeCodes", employeesTaskReportModel.EmployeeCodes);
                    cmd.Parameters.AddWithValue("@Activity", employeesTaskReportModel.Activity);
                    cmd.Parameters.AddWithValue("@FromDate", employeesTaskReportModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", employeesTaskReportModel.ToDate);

                    //Send Project Activities as xml
                    cmd.Parameters.Add(new SqlParameter("@EmployeeCodes", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    //Calling sp to get list of Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        EmployeesTaskReportModel employeeTaskReportModel = new EmployeesTaskReportModel();

                        employeeTaskReportModel.SlNo = SlNo;
                        employeeTaskReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                        employeeTaskReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                        employeeTaskReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                        employeeTaskReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                        employeeTaskReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                        employeeTaskReportModel.Department = sqlReader["Department"].ToString();
                        employeeTaskReportModel.Manager = sqlReader["Manager"].ToString();
                        employeeTaskReportModel.Activity = sqlReader["Activity"].ToString();
                        employeeTaskReportModel.ProductionAllocatedCount = Convert.ToInt64(sqlReader["ProductionAllocatedCount"]);
                        employeeTaskReportModel.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                        employeeTaskReportModel.QCAllocatedCount = Convert.ToInt64(sqlReader["QCAllocatedCount"]);
                        employeeTaskReportModel.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);
                        employeeTaskReportModel.ProductionHoursWorked = Convert.ToDecimal(sqlReader["ProductionHoursWorked"]);
                        employeeTaskReportModel.QCHoursWorked = Convert.ToDecimal(sqlReader["QCHoursWorked"]);
                        employeeTaskReportModel.NoOfProjects = Convert.ToInt32(sqlReader["NoOfProjects"]);
                        employeeTaskReportModel.InputCount = Convert.ToInt64(sqlReader["InputCount"]);
                        employeeTaskReportModel.ProductionTarget = Convert.ToInt32(sqlReader["ProductionTarget"]);
                        employeeTaskReportModel.QCTarget = Convert.ToInt32(sqlReader["QCTarget"]);
                        employeeTaskReportModel.ProductivityManDays = Convert.ToDecimal(sqlReader["ProductivityManDays"]);
                        employeeTaskReportModel.AveragePerDay = Convert.ToDecimal(sqlReader["AveragePerDay"]);
                        employeeTaskReportModel.ProductionAllocatedOn = sqlReader["ProductionAllocatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionAllocatedOn"];
                        employeeTaskReportModel.QCAllocatedOn = sqlReader["QCAllocatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["QCAllocatedOn"];
                        employeeTaskReportModel.ProductionStartDate = sqlReader["ProductionStartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionStartDate"];
                        employeeTaskReportModel.ProductionEndDate = sqlReader["ProductionEndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionEndDate"];
                        employeeTaskReportModel.QCStartDate = sqlReader["QCStartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["QCStartDate"];
                        employeeTaskReportModel.QCEndDate = sqlReader["QCEndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["QCEndDate"];
                        
                        employeesTaskReportDataList.Add(employeeTaskReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(employeesTaskReportDataList);
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

        #region Read Employees Task Summary Report data
        [HttpPost]
        [Route("ReadEmployeesTaskSummaryReportData")]
        public IHttpActionResult ReadEmployeesTaskSummaryReportData([FromBody] EmployeesTaskSummaryReportModel employeesTaskSummaryReportModel)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the Employees Task Summary Report data
                List<EmployeesTaskSummaryReportModel> employeesTaskSummaryReportModelDataList = new List<EmployeesTaskSummaryReportModel>();

                #region Employee Codes
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<EmployeeCodes>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, employeesTaskSummaryReportModel.EmployeeCodes);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                #region Status Options
                //create xml serialized data
                var StatusSerializer = new XmlSerializer(typeof(List<StatusOptions>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var StatusStream = new StringWriter();

                //write serialized data to stream
                StatusSerializer.Serialize(StatusStream, employeesTaskSummaryReportModel.StatusOptions);

                //Read stream as xml string
                StringReader StatusTransactionXml = new StringReader(StatusStream.ToString());

                //Read xml string
                XmlTextReader StatusXmlReader = new XmlTextReader(StatusTransactionXml);

                //Convert xml string to sql xml
                SqlXml StatusSqlXml = new SqlXml(StatusXmlReader);
                #endregion

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spEmployeesTaskSummary";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@CustomerCode", employeesTaskSummaryReportModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", employeesTaskSummaryReportModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", employeesTaskSummaryReportModel.BatchNo);
                    cmd.Parameters.AddWithValue("@Department", employeesTaskSummaryReportModel.Department);
                    cmd.Parameters.Add(new SqlParameter("@EmployeeCodes", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@Activity", employeesTaskSummaryReportModel.Department);
                    cmd.Parameters.AddWithValue("@FromDate", employeesTaskSummaryReportModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", employeesTaskSummaryReportModel.ToDate);
                    cmd.Parameters.Add(new SqlParameter("@StatusOptions", SqlDbType.Xml)
                    {
                        Value = StatusSqlXml
                    });

                     //Calling sp to get list of Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        EmployeesTaskSummaryReportModel employeeTaskSummaryReportModel = new EmployeesTaskSummaryReportModel();

                        employeeTaskSummaryReportModel.SlNo = SlNo;
                        employeeTaskSummaryReportModel.EmployeeNameCode = sqlReader["EmployeeName"].ToString();
                        employeeTaskSummaryReportModel.Projects = Convert.ToInt32(sqlReader["Projects"]);
                        employeeTaskSummaryReportModel.Activities = Convert.ToInt32(sqlReader["Activities"]);
                        employeeTaskSummaryReportModel.ProductionAllocatedCount = Convert.ToInt64(sqlReader["ProductionAllocatedCount"]);
                        employeeTaskSummaryReportModel.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                        employeeTaskSummaryReportModel.QCAllocatedCount = Convert.ToInt64(sqlReader["QCAllocatedCount"]);
                        employeeTaskSummaryReportModel.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);
                        employeeTaskSummaryReportModel.HoursWorked = Convert.ToDecimal(sqlReader["HoursWorked"]);
                        employeeTaskSummaryReportModel.ManDays = Convert.ToDecimal(sqlReader["ManDays"]);
                        employeeTaskSummaryReportModel.Status = sqlReader["Status"].ToString();
                        employeeTaskSummaryReportModel.ProductionAllocatedManDays = Convert.ToDecimal(sqlReader["ProductionAllocatedManDays"]);
                        employeeTaskSummaryReportModel.QCAllocatedManDays = Convert.ToDecimal(sqlReader["QCAllocatedManDays"]);

                        employeesTaskSummaryReportModelDataList.Add(employeeTaskSummaryReportModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(employeesTaskSummaryReportModelDataList);
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

        #region Export Employees Task Details Report data to Excel
        [HttpPost]
        [Route("ExportEmployeesTaskDetailsReportDataToExcel")]
        public HttpResponseMessage ExportEmployeesTaskDetailsReportDataToExcel([FromBody] EmployeesTaskReportModel employeesTaskReportModel)
        {
            try
            {
                if (employeesTaskReportModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Employees Task Details Report", employeesTaskReportModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "Employees Task Details Report.xlsx";

                    //Create a list to hold the Employees Task Details Report data
                    List<EmployeesTaskReportModel> employeesTaskReportDataList = new List<EmployeesTaskReportModel>();

                    System.Data.Common.DbDataReader sqlReader;
                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Employee Codes
                    //create xml serialized data
                    var serializer = new XmlSerializer(typeof(List<EmployeeCodes>),
                                           new XmlRootAttribute("root"));
                    //create a stream
                    var stream = new StringWriter();

                    //write serialized data to stream
                    serializer.Serialize(stream, employeesTaskReportModel.EmployeeCodes);

                    //Read stream as xml string
                    StringReader transactionXml = new StringReader(stream.ToString());

                    //Read xml string
                    XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                    //Convert xml string to sql xml
                    SqlXml sqlXml = new SqlXml(xmlReader);
                    #endregion

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
                        cmd.CommandText = "spEmployeesTaskReport";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", employeesTaskReportModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", employeesTaskReportModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", employeesTaskReportModel.BatchNo);
                        cmd.Parameters.AddWithValue("@Department", employeesTaskReportModel.Department);
                        //cmd.Parameters.AddWithValue("@EmployeeNameCode", employeesTaskReportModel.EmployeeName);
                        cmd.Parameters.AddWithValue("@Activity", employeesTaskReportModel.Activity);
                        cmd.Parameters.AddWithValue("@FromDate", employeesTaskReportModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", employeesTaskReportModel.ToDate);
                        //Send Project Activities as xml
                        cmd.Parameters.Add(new SqlParameter("@EmployeeCodes", SqlDbType.Xml)
                        {
                            Value = sqlXml
                        });

                        //Calling sp to get list of Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                EmployeesTaskReportModel employeeTaskReportModel = new EmployeesTaskReportModel();

                                employeeTaskReportModel.SlNo = SlNo;
                                employeeTaskReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                                employeeTaskReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                                employeeTaskReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                                employeeTaskReportModel.EmployeeCode = sqlReader["EmployeeCode"].ToString();
                                employeeTaskReportModel.EmployeeName = sqlReader["EmployeeName"].ToString();
                                employeeTaskReportModel.Department = sqlReader["Department"].ToString();
                                employeeTaskReportModel.Manager = sqlReader["Manager"].ToString();
                                employeeTaskReportModel.Activity = sqlReader["Activity"].ToString();
                                employeeTaskReportModel.ProductionAllocatedCount = Convert.ToInt32(sqlReader["ProductionAllocatedCount"]);
                                employeeTaskReportModel.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                employeeTaskReportModel.QCAllocatedCount = Convert.ToInt32(sqlReader["QCAllocatedCount"]);
                                employeeTaskReportModel.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                                employeeTaskReportModel.ProductionHoursWorked = Convert.ToDecimal(sqlReader["ProductionHoursWorked"]);
                                employeeTaskReportModel.QCHoursWorked = Convert.ToDecimal(sqlReader["QCHoursWorked"]);
                                employeeTaskReportModel.AveragePerDay = Convert.ToDecimal(sqlReader["AveragePerDay"]);
                                employeeTaskReportModel.ProductivityManDays = Convert.ToDecimal(sqlReader["ProductivityManDays"]);
                                employeeTaskReportModel.ProductionAllocatedOn = sqlReader["ProductionAllocatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionAllocatedOn"];
                                employeeTaskReportModel.QCAllocatedOn = sqlReader["QCAllocatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["QCAllocatedOn"];
                                employeeTaskReportModel.ProductionStartDate = sqlReader["ProductionStartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionStartDate"];
                                employeeTaskReportModel.ProductionEndDate = sqlReader["ProductionEndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["ProductionEndDate"];
                                employeeTaskReportModel.QCStartDate = sqlReader["QCStartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["QCStartDate"];
                                employeeTaskReportModel.QCEndDate = sqlReader["QCEndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["QCEndDate"];

                                employeesTaskReportDataList.Add(employeeTaskReportModel);
                                SlNo++;
                            }
                            conn.Close();
                            #endregion

                            #region Writing report header
                            for (int col = 0; col <= 22; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 23);
                            ws.Cells[1, 0].PutValue("Employees Task Report");

                            ws.Cells.Merge(3, 0, 1, 2);
                            ws.Cells[3, 0].PutValue("Customer Code : " + employeesTaskReportModel.CustomerCode);
                            ws.Cells.Merge(3, 2, 1, 2);
                            ws.Cells[3, 2].PutValue("Project Code : " + employeesTaskReportModel.ProjectCode);
                            if (employeesTaskReportModel.BatchNo.Trim().Length > 0)
                            {
                                ws.Cells.Merge(3, 4, 1, 2);
                                ws.Cells[3, 4].PutValue("Batch No. : " + employeesTaskReportModel.BatchNo);
                            }
                            ws.Cells.Merge(3, 6, 1, 2);
                            ws.Cells[3, 6].PutValue("Department : " + employeesTaskReportModel.Department);
                            ws.Cells.Merge(3, 8, 1, 2);
                            ws.Cells[3, 8].PutValue("Employee : " + employeesTaskReportModel.EmployeeName);
                            ws.Cells.Merge(3, 10, 1, 2);
                            ws.Cells[3, 10].PutValue("Activity : " + employeesTaskReportModel.Activity);
                            #endregion

                            #region Writing column headings
                            ws.Cells[5, 0].PutValue("S.No.");
                            ws.Cells[5, 1].PutValue("Customer Code");
                            ws.Cells[5, 2].PutValue("Project Code");
                            ws.Cells[5, 3].PutValue("Batch No.");
                            ws.Cells[5, 4].PutValue("Employee Code");
                            ws.Cells[5, 5].PutValue("Employee Name");
                            ws.Cells[5, 6].PutValue("Department");
                            ws.Cells[5, 7].PutValue("Manager");
                            ws.Cells[5, 8].PutValue("Activity");
                            ws.Cells[5, 9].PutValue("Production Allocated Count");
                            ws.Cells[5, 10].PutValue("Production Completed Count");
                            ws.Cells[5, 11].PutValue("QC Allocated Count");
                            ws.Cells[5, 12].PutValue("QC Completed Count");
                            ws.Cells[5, 13].PutValue("Production Hours Worked");
                            ws.Cells[5, 14].PutValue("QC Hours Worked");
                            ws.Cells[5, 15].PutValue("Average Per Day");
                            ws.Cells[5, 16].PutValue("Productivity (Man Days)");
                            ws.Cells[5, 17].PutValue("Production Allocated On");
                            ws.Cells[5, 18].PutValue("QC Allocated On");
                            ws.Cells[5, 19].PutValue("Production Start Date");
                            ws.Cells[5, 20].PutValue("Production End Date");
                            ws.Cells[5, 21].PutValue("QC Start Date");
                            ws.Cells[5, 22].PutValue("QC End Date");

                            for (int c = 0; c <= 22; c++)
                                ws.Cells[5, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (EmployeesTaskReportModel employeeTaskReportModel in employeesTaskReportDataList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(employeeTaskReportModel.SlNo);
                                ws.Cells[row, 1].PutValue(employeeTaskReportModel.CustomerCode);
                                ws.Cells[row, 2].PutValue(employeeTaskReportModel.ProjectCode);
                                if (!string.IsNullOrEmpty(employeeTaskReportModel.BatchNo))
                                    ws.Cells[row, 3].PutValue(employeeTaskReportModel.BatchNo);
                                ws.Cells[row, 4].PutValue(employeeTaskReportModel.EmployeeCode);
                                ws.Cells[row, 5].PutValue(employeeTaskReportModel.EmployeeName);
                                if (!string.IsNullOrEmpty(employeeTaskReportModel.Department))
                                    ws.Cells[row, 6].PutValue(employeeTaskReportModel.Department);
                                if (!string.IsNullOrEmpty(employeeTaskReportModel.Manager))
                                    ws.Cells[row, 7].PutValue(employeeTaskReportModel.Manager);
                                ws.Cells[row, 8].PutValue(employeeTaskReportModel.Activity);
                                ws.Cells[row, 9].PutValue(employeeTaskReportModel.ProductionAllocatedCount);
                                ws.Cells[row, 10].PutValue(employeeTaskReportModel.ProductionCompletedCount);
                                ws.Cells[row, 11].PutValue(employeeTaskReportModel.QCAllocatedCount);
                                ws.Cells[row, 12].PutValue(employeeTaskReportModel.QCCompletedCount);
                                ws.Cells[row, 13].PutValue(employeeTaskReportModel.ProductionHoursWorked);
                                ws.Cells[row, 14].PutValue(employeeTaskReportModel.QCHoursWorked);
                                ws.Cells[row, 15].PutValue(employeeTaskReportModel.AveragePerDay);
                                ws.Cells[row, 16].PutValue(employeeTaskReportModel.ProductivityManDays);
                                if (employeeTaskReportModel.ProductionAllocatedOn != null)
                                    ws.Cells[row, 17].PutValue(Convert.ToDateTime(employeeTaskReportModel.ProductionAllocatedOn).ToString("dd-MMM-yyyy"));
                                if (employeeTaskReportModel.QCAllocatedOn != null)
                                    ws.Cells[row, 18].PutValue(Convert.ToDateTime(employeeTaskReportModel.QCAllocatedOn).ToString("dd-MMM-yyyy"));
                                if (employeeTaskReportModel.ProductionStartDate != null)
                                    ws.Cells[row, 19].PutValue(Convert.ToDateTime(employeeTaskReportModel.ProductionStartDate).ToString("dd-MMM-yyyy"));
                                if (employeeTaskReportModel.ProductionEndDate != null)
                                    ws.Cells[row, 20].PutValue(Convert.ToDateTime(employeeTaskReportModel.ProductionEndDate).ToString("dd-MMM-yyyy"));
                                if (employeeTaskReportModel.QCStartDate != null)
                                    ws.Cells[row, 21].PutValue(Convert.ToDateTime(employeeTaskReportModel.QCStartDate).ToString("dd-MMM-yyyy"));
                                if (employeeTaskReportModel.QCEndDate != null)
                                    ws.Cells[row, 22].PutValue(Convert.ToDateTime(employeeTaskReportModel.QCEndDate).ToString("dd-MMM-yyyy"));

                                #endregion

                                #region Setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 12].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 13].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 16].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 19].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 20].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 21].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 22].SetStyle(styleCenterAlignData);
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

        #region Export Employees Task Summary Report data to Excel
        [HttpPost]
        [Route("ExportEmployeesTaskSummaryReportDataToExcel")]
        public HttpResponseMessage ExportEmployeesTaskSummaryReportDataToExcel([FromBody] EmployeesTaskSummaryReportModel employeesTaskSummaryReportModel)
        {
            try
            {
                if (employeesTaskSummaryReportModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Employees Task Summary Report", employeesTaskSummaryReportModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "Employees Task Summary Report.xlsx";

                    #region Employee Codes
                    //create xml serialized data
                    var serializer = new XmlSerializer(typeof(List<EmployeeCodes>),
                                           new XmlRootAttribute("root"));
                    //create a stream
                    var stream = new StringWriter();

                    //write serialized data to stream
                    serializer.Serialize(stream, employeesTaskSummaryReportModel.EmployeeCodes);

                    //Read stream as xml string
                    StringReader transactionXml = new StringReader(stream.ToString());

                    //Read xml string
                    XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                    //Convert xml string to sql xml
                    SqlXml sqlXml = new SqlXml(xmlReader);
                    #endregion

                    #region Status Options
                    //create xml serialized data
                    var StatusSerializer = new XmlSerializer(typeof(List<StatusOptions>),
                                           new XmlRootAttribute("root"));
                    //create a stream
                    var StatusStream = new StringWriter();

                    //write serialized data to stream
                    StatusSerializer.Serialize(StatusStream, employeesTaskSummaryReportModel.StatusOptions);

                    //Read stream as xml string
                    StringReader StatusTransactionXml = new StringReader(StatusStream.ToString());

                    //Read xml string
                    XmlTextReader StatusXmlReader = new XmlTextReader(StatusTransactionXml);

                    //Convert xml string to sql xml
                    SqlXml StatusSqlXml = new SqlXml(StatusXmlReader);
                    #endregion

                    //Create a list to hold the Employees Task Summary Report data
                    List<EmployeesTaskSummaryReportModel> employeesTaskSummaryReportModelDataList = new List<EmployeesTaskSummaryReportModel>();

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
                        cmd.CommandText = "spEmployeesTaskSummary";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@CustomerCode", employeesTaskSummaryReportModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", employeesTaskSummaryReportModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", employeesTaskSummaryReportModel.BatchNo);
                        cmd.Parameters.AddWithValue("@Department", employeesTaskSummaryReportModel.Department);
                        cmd.Parameters.Add(new SqlParameter("@EmployeeCodes", SqlDbType.Xml)
                        {
                            Value = sqlXml
                        });
                        cmd.Parameters.AddWithValue("@Activity", employeesTaskSummaryReportModel.Department);
                        cmd.Parameters.AddWithValue("@FromDate", employeesTaskSummaryReportModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", employeesTaskSummaryReportModel.ToDate);
                        cmd.Parameters.Add(new SqlParameter("@StatusOptions", SqlDbType.Xml)
                        {
                            Value = StatusSqlXml
                        });

                        //Calling sp to get list of Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                EmployeesTaskSummaryReportModel employeeTaskSummaryReportModel = new EmployeesTaskSummaryReportModel();

                                employeeTaskSummaryReportModel.SlNo = SlNo;
                                employeeTaskSummaryReportModel.EmployeeNameCode = sqlReader["EmployeeName"].ToString();
                                employeeTaskSummaryReportModel.Projects = Convert.ToInt32(sqlReader["Projects"]);
                                employeeTaskSummaryReportModel.Activities = Convert.ToInt32(sqlReader["Activities"]);
                                employeeTaskSummaryReportModel.ProductionAllocatedCount = Convert.ToInt64(sqlReader["ProductionAllocatedCount"]);
                                employeeTaskSummaryReportModel.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                                employeeTaskSummaryReportModel.QCAllocatedCount = Convert.ToInt64(sqlReader["QCAllocatedCount"]);
                                employeeTaskSummaryReportModel.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);
                                employeeTaskSummaryReportModel.HoursWorked = Convert.ToDecimal(sqlReader["HoursWorked"]);
                                employeeTaskSummaryReportModel.ManDays = Convert.ToDecimal(sqlReader["ManDays"]);
                                employeeTaskSummaryReportModel.Status = sqlReader["Status"].ToString();
                                employeeTaskSummaryReportModel.ProductionAllocatedManDays = Convert.ToDecimal(sqlReader["ProductionAllocatedManDays"]);
                                employeeTaskSummaryReportModel.QCAllocatedManDays = Convert.ToDecimal(sqlReader["QCAllocatedManDays"]);

                                employeesTaskSummaryReportModelDataList.Add(employeeTaskSummaryReportModel);
                                SlNo++;
                            }
                            conn.Close();
                            #endregion

                            #region Writing report header
                            for (int col = 0; col <= 12; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 13);
                            ws.Cells[1, 0].PutValue("Employees Task Summary Report");
                            #endregion

                            #region Writing column headings
                            ws.Cells[3, 0].PutValue("S.No.");
                            ws.Cells[3, 1].PutValue("Employee Name");
                            ws.Cells[3, 2].PutValue("Projects");
                            ws.Cells[3, 3].PutValue("Activities");
                            ws.Cells[3, 4].PutValue("Production Allocated Count");
                            ws.Cells[3, 5].PutValue("Production Completed Count");
                            ws.Cells[3, 6].PutValue("QC Allocated Count");
                            ws.Cells[3, 7].PutValue("QC Completed Count");
                            ws.Cells[3, 8].PutValue("Hours Worked");
                            ws.Cells[3, 9].PutValue("Man Days");
                            ws.Cells[3, 10].PutValue("Status");
                            ws.Cells[3, 11].PutValue("Production Allocated Man Days");
                            ws.Cells[3, 12].PutValue("QC Allocated Man Days");

                            for (int c = 0; c <= 12; c++)
                                ws.Cells[3, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (EmployeesTaskSummaryReportModel employeeTaskReportModel in employeesTaskSummaryReportModelDataList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(employeeTaskReportModel.SlNo);
                                ws.Cells[row, 1].PutValue(employeeTaskReportModel.EmployeeNameCode);
                                ws.Cells[row, 2].PutValue(employeeTaskReportModel.Projects);
                                ws.Cells[row, 3].PutValue(employeeTaskReportModel.Activities);
                                ws.Cells[row, 4].PutValue(employeeTaskReportModel.ProductionAllocatedCount);
                                ws.Cells[row, 5].PutValue(employeeTaskReportModel.ProductionCompletedCount);
                                ws.Cells[row, 6].PutValue(employeeTaskReportModel.QCAllocatedCount);
                                ws.Cells[row, 7].PutValue(employeeTaskReportModel.QCCompletedCount);
                                ws.Cells[row, 8].PutValue(employeeTaskReportModel.HoursWorked);
                                ws.Cells[row, 9].PutValue(employeeTaskReportModel.ManDays);
                                ws.Cells[row, 10].PutValue(employeeTaskReportModel.Status);
                                ws.Cells[row, 11].PutValue(employeeTaskReportModel.ProductionAllocatedManDays);
                                ws.Cells[row, 12].PutValue(employeeTaskReportModel.QCAllocatedManDays);
                                #endregion

                                #region Setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 12].SetStyle(styleCenterAlignData);
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

        #region Read UnAllocated Resource Details Report data
        [HttpPost]
        [Route("ReadUnAllocatedResourceDetailsReportData")]
        public IHttpActionResult ReadUnAllocatedResourceDetailsReportData([FromBody] UnAllocatedResourceModel unAllocatedResourceModel)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the unAllocated Resource Details Report data
                List<UnAllocatedResourceModel> unAllocatedResourceDetailsList = new List<UnAllocatedResourceModel>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUnAllocatedResources";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@FromDate", unAllocatedResourceModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", unAllocatedResourceModel.ToDate);

                    //Calling sp to get list of Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        UnAllocatedResourceModel unAllocatedResourceDetailsModel = new UnAllocatedResourceModel();

                        unAllocatedResourceDetailsModel.SlNo = SlNo;
                        unAllocatedResourceDetailsModel.EmployeeCode = sqlReader["Username"].ToString();
                        unAllocatedResourceDetailsModel.EmployeeName = sqlReader["UserFullName"].ToString();
                        unAllocatedResourceDetailsModel.Department = sqlReader["Department"].ToString();
                        unAllocatedResourceDetailsModel.Manager = sqlReader["ManagerNameCode"].ToString();

                        unAllocatedResourceDetailsList.Add(unAllocatedResourceDetailsModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(unAllocatedResourceDetailsList);
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

        #region Export UnAllocated Resource Details Report data
        [HttpPost]
        [Route("ExportUnAllocatedResourceDetailsReportDataToExcel")]
        public HttpResponseMessage ExportUnAllocatedResourceDetailsReportDataToExcel([FromBody] UnAllocatedResourceModel unAllocatedResourceModel)
        {
            try
            {
                if (unAllocatedResourceModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Un-Allocated Resource Details Report", unAllocatedResourceModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "UnAllocated Resource Details Report.xlsx";

                    //Create a list to hold the unAllocated Resource Details Report data
                    List<UnAllocatedResourceModel> unAllocatedResourceDetailsList = new List<UnAllocatedResourceModel>();

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
                        cmd.CommandText = "spUnAllocatedResources";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@FromDate", unAllocatedResourceModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", unAllocatedResourceModel.ToDate);

                        //Calling sp to get list of Report Data
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                UnAllocatedResourceModel unAllocatedResourceDetailsModel = new UnAllocatedResourceModel();

                                unAllocatedResourceDetailsModel.SlNo = SlNo;
                                unAllocatedResourceDetailsModel.EmployeeCode = sqlReader["Username"].ToString();
                                unAllocatedResourceDetailsModel.EmployeeName = sqlReader["UserFullName"].ToString();
                                unAllocatedResourceDetailsModel.Department = sqlReader["Department"].ToString();
                                unAllocatedResourceDetailsModel.Manager = sqlReader["ManagerNameCode"].ToString();

                                unAllocatedResourceDetailsList.Add(unAllocatedResourceDetailsModel);
                                SlNo++;
                            }
                            conn.Close();
                            #endregion

                            #region Writing report header
                            for (int col = 0; col <= 4; col++)
                                ws.Cells[1, col].SetStyle(styleCenterAlignData);

                            ws.Cells.Merge(1, 0, 1, 5);
                            ws.Cells[1, 0].PutValue("Unallocated Resource Details Report");
                            #endregion

                            #region Writing column headings
                            ws.Cells[3, 0].PutValue("S.No.");
                            ws.Cells[3, 1].PutValue("Employee Code");
                            ws.Cells[3, 2].PutValue("Employee Name");
                            ws.Cells[3, 3].PutValue("Department");
                            ws.Cells[3, 4].PutValue("Manager");

                            for (int c = 0; c <= 4; c++)
                                ws.Cells[3, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (UnAllocatedResourceModel unAllocatedResourceDetailsModel in unAllocatedResourceDetailsList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(unAllocatedResourceDetailsModel.SlNo);
                                ws.Cells[row, 1].PutValue(unAllocatedResourceDetailsModel.EmployeeCode);
                                ws.Cells[row, 2].PutValue(unAllocatedResourceDetailsModel.EmployeeName);
                                ws.Cells[row, 3].PutValue(unAllocatedResourceDetailsModel.Department);
                                ws.Cells[row, 4].PutValue(unAllocatedResourceDetailsModel.Manager);
                                #endregion

                                #region Setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
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
