using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
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
    [RoutePrefix("api/Dashboard")]
    public class DashboardController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;

        public DashboardController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }


        #region Read Dashboard Project Status, Projects, Customers
        [HttpGet]
        [Route]
        public IHttpActionResult ReadDashboardDetails()
        {
            try
            {
                Dashboard dashboard = new Dashboard();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spDashboard";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get dashboard details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        dashboard.NoOfCustomers = Convert.ToInt32(sqlReader["NoOfCustomers"]);
                        dashboard.NoOfProjects = Convert.ToInt32(sqlReader["NoOfProjects"]);
                        dashboard.NoOfOnGoingProjects = Convert.ToInt32(sqlReader["NoOfOnGoingProjects"]);
                        dashboard.NoOfCompletedProjects = Convert.ToInt32(sqlReader["NoOfCompletedProjects"]);
                        dashboard.NoOfPendingProjects = Convert.ToInt32(sqlReader["NoOfPendingProjects"]);
                        dashboard.NoOfCompletedProjectsPercentage = Convert.ToDecimal(sqlReader["NoOfCompletedProjectsPercentage"]);
                        dashboard.NoOfPendingProjectsPercentage = Convert.ToDecimal(sqlReader["NoOfPendingProjectsPercentage"]);
                        dashboard.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);
                        dashboard.NoOfCompletedBatches = Convert.ToInt32(sqlReader["NoOfCompletedBatches"]);
                        dashboard.NoOfPendingBatches = Convert.ToInt32(sqlReader["NoOfPendingBatches"]);
                        dashboard.NoOfCompletedBatchesPercentage = Convert.ToDecimal(sqlReader["NoOfCompletedBatchesPercentage"]);
                        dashboard.NoOfPendingBatchesPercentage = Convert.ToDecimal(sqlReader["NoOfPendingBatchesPercentage"]);
                        dashboard.NoOfActiveTasks = Convert.ToInt32(sqlReader["NoOfActiveTasks"]);
                        dashboard.NoOfActiveResources = Convert.ToInt32(sqlReader["NoOfActiveResources"]);
                        conn.Close();
                        //return dashboard to the request
                        return Ok(dashboard);
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }
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

        #region On-Going Projects Received Within Period
        [HttpGet]
        [Route("ReadActiveProjects")]
        public IHttpActionResult ReadActiveProjects(DateTime FromDate, DateTime ToDate)
        {
            try
            {
                List<Project> activeProjectsList = new List<Project>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectsReceivedWithinPeriod";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get the Active Projects List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Project project = new Project();

                        project.ProjectID = Convert.ToInt32(sqlReader["ProjectID"]);
                        project.CustomerCode = sqlReader["CustomerCode"].ToString();
                        project.ProjectCode = sqlReader["ProjectCode"].ToString();
                        project.Scope = sqlReader["Scope"].ToString();
                        project.ProjectType = sqlReader["ProjectType"].ToString();
                        project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                        project.Status = sqlReader["Status"].ToString();

                        activeProjectsList.Add(project);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(activeProjectsList);
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

        #region Read Active Tasks
        [HttpGet]
        [Route("ReadActiveTasks")]
        public IHttpActionResult ReadActiveTasks()
        {
            try
            {
                int SlNo = 1;
                List<ActiveTasksDetails> activeTasksDetailsList = new List<ActiveTasksDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectActivitiesStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get the Active Tasks Details List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ActiveTasksDetails activeTasksDetails = new ActiveTasksDetails();

                        activeTasksDetails.SlNo = SlNo;
                        activeTasksDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        activeTasksDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        activeTasksDetails.BatchNo = sqlReader["BatchNo"].ToString();
                        activeTasksDetails.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        activeTasksDetails.Activity = sqlReader["Activity"].ToString();
                        activeTasksDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        activeTasksDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);

                        activeTasksDetailsList.Add(activeTasksDetails);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(activeTasksDetailsList);
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

        #region Read Active Resources
        [HttpGet]
        [Route("ReadResources")]
        public IHttpActionResult ReadResources(DateTime? FromDate, DateTime? ToDate)
        {
            try
            {
                int SlNo = 1;
                List<ResourceDetails> resourceDetailsList = new List<ResourceDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectsResources";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get list of resource Details List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ResourceDetails resourceDetails = new ResourceDetails();

                        resourceDetails.SlNo = SlNo;
                        resourceDetails.ResourceCode = sqlReader["UserID"].ToString();
                        resourceDetails.ResourceName = sqlReader["UserName"].ToString();
                        resourceDetails.StartDate = sqlReader["StartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["StartDate"];
                        resourceDetails.EndDate = sqlReader["EndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["EndDate"];
                        resourceDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        resourceDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        resourceDetails.HoursWorked = Convert.ToInt32(sqlReader["TotalHoursWorked"]);

                        resourceDetailsList.Add(resourceDetails);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(resourceDetailsList);
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

        #region Read Resource Productivity Details
        [HttpGet]
        [Route("ReadResourceProductivityDetails")]
        public IHttpActionResult ReadResourceProductivityDetails(string UserID, DateTime? FromDate, DateTime? ToDate)
        {
            try
            {
                int SlNo = 1;  //temporary added since bootstrap table requires a key to render table
                List<ResourceProductivityDetails> resourceProductivityDetailsList = new List<ResourceProductivityDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spResourceProductivityDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get list of resource productivity Details List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ResourceProductivityDetails resourceProductivityDetails = new ResourceProductivityDetails();

                        resourceProductivityDetails.SlNo = SlNo;
                        resourceProductivityDetails.DateWorked = (DateTime)sqlReader["DateWorked"];
                        resourceProductivityDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        resourceProductivityDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        resourceProductivityDetails.Activity = sqlReader["Activity"].ToString();
                        resourceProductivityDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        resourceProductivityDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        resourceProductivityDetails.ProductionTarget = Convert.ToInt32(sqlReader["ProductionTarget"]);
                        resourceProductivityDetails.QCTarget = Convert.ToInt32(sqlReader["QCTarget"]);
                        resourceProductivityDetails.HoursWorked = Convert.ToInt32(sqlReader["HoursWorked"]);

                        resourceProductivityDetailsList.Add(resourceProductivityDetails);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(resourceProductivityDetailsList);
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

        #region Read No.of Hours Worked
        [HttpGet]
        [Route("ReadHoursWorked")]
        public IHttpActionResult ReadHoursWorked()
        {
            try
            {
                int SlNo = 1;  
                List<HoursWorkedDetails> hoursWorkedDetailsList = new List<HoursWorkedDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectsHoursWorked";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get the worked hours details List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        HoursWorkedDetails hoursWorkedDetails = new HoursWorkedDetails();

                        hoursWorkedDetails.SlNo = SlNo;
                        hoursWorkedDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        hoursWorkedDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        hoursWorkedDetails.BatchNo = sqlReader["BatchNo"].ToString();
                        hoursWorkedDetails.ResourceCode = sqlReader["EmployeeCode"].ToString();
                        hoursWorkedDetails.ResourceName = sqlReader["EmployeeName"].ToString();
                        hoursWorkedDetails.NoOfHoursWorked = Convert.ToInt32(sqlReader["NoOfHoursWorked"]);

                        hoursWorkedDetailsList.Add(hoursWorkedDetails);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(hoursWorkedDetailsList);
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

        #region Projects Completion Status (for Chart)
        [HttpGet]
        [Route("ReadProjectsCompletionStatus")]
        public IHttpActionResult ReadProjectsCompletionStatus()
        {
            try
            {
                List<ProjectsCompletionStatus> projectsCompletionStatusList = new List<ProjectsCompletionStatus>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectsCompletionStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get the Projects Completion Status List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectsCompletionStatus projectsCompletionStatus = new ProjectsCompletionStatus();
                        projectsCompletionStatus.Status = sqlReader["Status"].ToString();
                        projectsCompletionStatus.NoOfProjects = Convert.ToInt32(sqlReader["NoOfProjects"]);
                        projectsCompletionStatusList.Add(projectsCompletionStatus);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectsCompletionStatusList);
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

        #region Export To Excel
        #region Export On-Going Projects Received Within Period To Excel
        [HttpGet]
        [Route("ExportActiveProjectsToExcel")]
        public HttpResponseMessage ExportActiveProjectsToExcel(DateTime FromDate, DateTime ToDate)
        {
            try
            {
                string FileName = "Active Projects.xlsx";

                //Create a list to hold the list of Active Projects
                List<Project> activeProjectsList = new List<Project>();
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
                int row = 1;
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
                    cmd.CommandText = "spOnGoingProjectsReceivedWithinPeriod";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Calling sp to get the Active Projects List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            Project project = new Project();

                            project.ProjectID = Convert.ToInt32(sqlReader["ProjectID"]);
                            project.CustomerCode = sqlReader["CustomerCode"].ToString();
                            project.ProjectCode = sqlReader["ProjectCode"].ToString();
                            project.Scope = sqlReader["Scope"].ToString();
                            project.ProjectType = sqlReader["ProjectType"].ToString();
                            project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                            project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                            project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                            project.Status = sqlReader["Status"].ToString();

                            activeProjectsList.Add(project);
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Customer Code");
                        ws.Cells[0, 2].PutValue("Project Code");
                        ws.Cells[0, 3].PutValue("Scope");
                        ws.Cells[0, 4].PutValue("Project Type");
                        ws.Cells[0, 5].PutValue("Received Date");
                        ws.Cells[0, 6].PutValue("Input Count");
                        ws.Cells[0, 7].PutValue("Planned Delivery Date");
                        ws.Cells[0, 8].PutValue("Status");

                        for (int c = 0; c <= 8; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (Project project in activeProjectsList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(project.CustomerCode);
                            ws.Cells[row, 2].PutValue(project.ProjectCode);
                            ws.Cells[row, 3].PutValue(project.Scope);
                            ws.Cells[row, 4].PutValue(project.ProjectType);
                            ws.Cells[row, 5].PutValue(Convert.ToDateTime(project.ReceivedDate).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 6].PutValue(project.InputCount);
                            if (project.PlannedDeliveryDate != null)
                                ws.Cells[row, 7].PutValue(Convert.ToDateTime(project.PlannedDeliveryDate).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 8].PutValue(project.Status);
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 8].SetStyle(styleCenterAlignData);
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
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Export Active Tasks To Excel
        [HttpGet]
        [Route("ExportActiveTasksToExcel")]
        public HttpResponseMessage ExportActiveTasksToExcel(string UserID)
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Active Tasks Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = "Active Tasks.xlsx";

                    //Create a list to hold the list of Active Tasks
                    List<ActiveTasksDetails> activeTasksList = new List<ActiveTasksDetails>();
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
                    int row = 1;
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
                        cmd.CommandText = "spOnGoingProjectActivitiesStatus";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Calling sp to get the Active Tasks Details List
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            while (sqlReader.Read())
                            {
                                ActiveTasksDetails activeTasksDetails = new ActiveTasksDetails();

                                activeTasksDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                                activeTasksDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                                activeTasksDetails.BatchNo = sqlReader["BatchNo"].ToString();
                                activeTasksDetails.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                                activeTasksDetails.Activity = sqlReader["Activity"].ToString();
                                activeTasksDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                activeTasksDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);

                                activeTasksList.Add(activeTasksDetails);
                            }
                            conn.Close();

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Customer Code");
                            ws.Cells[0, 2].PutValue("Project Code");
                            ws.Cells[0, 3].PutValue("Batch No.");
                            ws.Cells[0, 4].PutValue("Input Count");
                            ws.Cells[0, 5].PutValue("Activity / Task");
                            ws.Cells[0, 6].PutValue("Production Completed Count");
                            ws.Cells[0, 7].PutValue("QC Completed Count");

                            for (int c = 0; c <= 7; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (ActiveTasksDetails activeTasksDetails in activeTasksList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(row);
                                ws.Cells[row, 1].PutValue(activeTasksDetails.CustomerCode);
                                ws.Cells[row, 2].PutValue(activeTasksDetails.ProjectCode);
                                ws.Cells[row, 3].PutValue(activeTasksDetails.BatchNo);
                                ws.Cells[row, 4].PutValue(activeTasksDetails.InputCount);
                                ws.Cells[row, 5].PutValue(activeTasksDetails.Activity);
                                ws.Cells[row, 6].PutValue(activeTasksDetails.ProductionCompletedCount);
                                ws.Cells[row, 7].PutValue(activeTasksDetails.QCCompletedCount);
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

        #region Export Active Resources To Excel
        [HttpGet]
        [Route("ExportActiveResourcesToExcel")]
        public HttpResponseMessage ExportActiveResourcesToExcel(DateTime? FromDate, DateTime? ToDate, string UserID)
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Active Resources Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    int SlNo = 1;
                    string FileName = "Active Resources.xlsx";

                    //Create a list to hold the list of Active Resources
                    List<ResourceDetails> resourceDetailsList = new List<ResourceDetails>();
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
                    int row = 1;
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
                        cmd.CommandText = "spOnGoingProjectsResources";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);

                        //Calling sp to get list of resource Details List
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            while (sqlReader.Read())
                            {
                                ResourceDetails resourceDetails = new ResourceDetails();

                                resourceDetails.SlNo = SlNo;
                                resourceDetails.ResourceCode = sqlReader["UserID"].ToString();
                                resourceDetails.ResourceName = sqlReader["UserName"].ToString();
                                resourceDetails.StartDate = sqlReader["StartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["StartDate"];
                                resourceDetails.EndDate = sqlReader["EndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["EndDate"];
                                resourceDetails.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                                resourceDetails.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                                resourceDetails.HoursWorked = Convert.ToInt32(sqlReader["TotalHoursWorked"]);

                                resourceDetailsList.Add(resourceDetails);
                                SlNo++;
                            }
                            conn.Close();

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Resource Code");
                            ws.Cells[0, 2].PutValue("Resource Name");
                            ws.Cells[0, 3].PutValue("Start Date");
                            ws.Cells[0, 4].PutValue("End Date");
                            ws.Cells[0, 5].PutValue("Production Completed Count");
                            ws.Cells[0, 6].PutValue("QC Completed Count");
                            ws.Cells[0, 7].PutValue("Hours Worked");

                            for (int c = 0; c <= 7; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (ResourceDetails resourceDetails in resourceDetailsList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(resourceDetails.SlNo);
                                ws.Cells[row, 1].PutValue(resourceDetails.ResourceCode);
                                ws.Cells[row, 2].PutValue(resourceDetails.ResourceName);
                                if (resourceDetails.StartDate != null)
                                    ws.Cells[row, 3].PutValue(Convert.ToDateTime(resourceDetails.StartDate).ToString("dd-MMM-yyyy"));
                                if (resourceDetails.EndDate != null)
                                    ws.Cells[row, 4].PutValue(Convert.ToDateTime(resourceDetails.EndDate).ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 5].PutValue(resourceDetails.ProductionCompletedCount);
                                ws.Cells[row, 6].PutValue(resourceDetails.QCCompletedCount);
                                ws.Cells[row, 7].PutValue(resourceDetails.HoursWorked);
                                #endregion

                                #region setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 7].SetStyle(styleCenterAlignData);
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

        #region Export No. of Hours Worked (Duration) To Excel
        [HttpGet]
        [Route("ExportNoOfHoursWorkedToExcel")]
        public HttpResponseMessage ExportNoOfHoursWorkedToExcel(string UserID)
        {
            try
            {
                if (UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("Duration / Hours Worked Report", UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = "Hours Worked.xlsx";

                    //Create a list to hold the list of hours worked
                    List<HoursWorkedDetails> hoursWorkedDetailsList = new List<HoursWorkedDetails>();
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
                    int row = 1;
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
                        cmd.CommandText = "spOnGoingProjectsHoursWorked";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        //Calling sp to get the worked hours details List
                        conn.Open();
                        cmd.CommandTimeout = 0;
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            while (sqlReader.Read())
                            {
                                HoursWorkedDetails hoursWorkedDetails = new HoursWorkedDetails();

                                hoursWorkedDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                                hoursWorkedDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                                hoursWorkedDetails.BatchNo = sqlReader["BatchNo"].ToString();
                                hoursWorkedDetails.ResourceCode = sqlReader["EmployeeCode"].ToString();
                                hoursWorkedDetails.ResourceName = sqlReader["EmployeeName"].ToString();
                                hoursWorkedDetails.NoOfHoursWorked = Convert.ToInt32(sqlReader["NoOfHoursWorked"]);

                                hoursWorkedDetailsList.Add(hoursWorkedDetails);
                            }
                            conn.Close();

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Customer Code");
                            ws.Cells[0, 2].PutValue("Project Code");
                            ws.Cells[0, 3].PutValue("Batch No.");
                            ws.Cells[0, 4].PutValue("Employee Code");
                            ws.Cells[0, 5].PutValue("Employee Name");
                            ws.Cells[0, 6].PutValue("No. of Hours Worked");

                            for (int c = 0; c <= 6; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (HoursWorkedDetails hoursWorkedDetails in hoursWorkedDetailsList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(row);
                                ws.Cells[row, 1].PutValue(hoursWorkedDetails.CustomerCode);
                                ws.Cells[row, 2].PutValue(hoursWorkedDetails.ProjectCode);
                                ws.Cells[row, 3].PutValue(hoursWorkedDetails.BatchNo);
                                ws.Cells[row, 4].PutValue(hoursWorkedDetails.ResourceCode);
                                ws.Cells[row, 5].PutValue(hoursWorkedDetails.ResourceName);
                                ws.Cells[row, 6].PutValue(hoursWorkedDetails.NoOfHoursWorked);
                                #endregion

                                #region setting row data style
                                ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                                ws.Cells[row, 5].SetStyle(styleLeftAlignData);
                                ws.Cells[row, 6].SetStyle(styleCenterAlignData);
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
        #endregion
    }
}
