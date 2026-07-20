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
    [RoutePrefix("api/project")]
    public class ProjectController : ApiController
    {
        private BLLAccessControl _BLLAccessControl;
        public ProjectController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLAccessControl = new BLLAccessControl(connectionString);
        }

        #region Create Project
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateProject([FromBody]Project project)
        {
            try
            {
                string CustomerInputFileExtension = string.Empty;
                string DeliveryPlanFileExtension = string.Empty;
                string ScopeFileExtension = string.Empty;
                string GuidelineFileExtension = string.Empty;
                string ChecklistFileExtension = string.Empty;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Data");

                if (!AccessControl.CanUserAccessPage(project.UserID, "Create Project"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Get all the files extension from temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                //Customer Input File
                if (File.Exists(dirTemp + project.CustomerInputFileName))
                    CustomerInputFileExtension = Path.GetExtension(dirTemp + project.CustomerInputFileName);

                //Delivery Plan File
                if (File.Exists(dirTemp + project.DeliveryPlanFileName))
                    DeliveryPlanFileExtension = Path.GetExtension(dirTemp + project.DeliveryPlanFileName);

                //Scope File
                if (File.Exists(dirTemp + project.ScopeFileName))
                    ScopeFileExtension = Path.GetExtension(dirTemp + project.ScopeFileName);

                //Guideline File
                if (File.Exists(dirTemp + project.GuidelineFileName))
                    GuidelineFileExtension = Path.GetExtension(dirTemp + project.GuidelineFileName);

                //Checklist File
                if (File.Exists(dirTemp + project.ChecklistFileName))
                    ChecklistFileExtension = Path.GetExtension(dirTemp + project.ChecklistFileName);
                #endregion

                #region Project Activities
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ActivityTarget>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, project.Activities);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCreateProject";

                    #region Adding Stored Procedure Parameters
                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", project.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectType", project.ProjectType);
                    cmd.Parameters.AddWithValue("@LocationCode", project.LocationCode);
                    cmd.Parameters.AddWithValue("@InputCount", project.InputCount);
                    cmd.Parameters.AddWithValue("@InputCountType", project.InputCountType);
                    cmd.Parameters.AddWithValue("@TypeOfInput", project.TypeOfInput);
                    cmd.Parameters.AddWithValue("@IsResourceBased", project.IsResourceBased);
                    cmd.Parameters.AddWithValue("@PlannedStartDate", project.PlannedStartDate);
                    cmd.Parameters.AddWithValue("@Remarks", project.Remarks);
                    cmd.Parameters.AddWithValue("@CustomerInputFileExtension", CustomerInputFileExtension);

                    //Project Received Details
                    cmd.Parameters.AddWithValue("@ReceivedDate", project.ReceivedDate);
                    cmd.Parameters.AddWithValue("@ReceivedFormat", project.ReceivedFormat);

                    //Project Delivery Details
                    cmd.Parameters.AddWithValue("@OutputFormat", project.OutputFormat);
                    cmd.Parameters.AddWithValue("@DeliveryMode", project.DeliveryMode);
                    cmd.Parameters.AddWithValue("@PlannedDeliveryDate", project.PlannedDeliveryDate);
                    cmd.Parameters.AddWithValue("@DeliveryPlanFileExtension", DeliveryPlanFileExtension);

                    //Project Scope
                    cmd.Parameters.AddWithValue("@Scope", project.Scope);
                    cmd.Parameters.AddWithValue("@ScopeFileExtension", ScopeFileExtension);

                    //Project Guideline
                    cmd.Parameters.AddWithValue("@Guideline", project.Guideline);
                    cmd.Parameters.AddWithValue("@GuidelineFileExtension", GuidelineFileExtension);

                    //Project Checklist
                    cmd.Parameters.AddWithValue("@Checklist", project.Checklist);
                    cmd.Parameters.AddWithValue("@ChecklistFileExtension", ChecklistFileExtension);

                    //Project Client Details
                    cmd.Parameters.AddWithValue("@EmailDate", project.EmailDate);
                    cmd.Parameters.AddWithValue("@EmailDescription", project.EmailDescription);

                    //Project UNSPSC Version
                    cmd.Parameters.AddWithValue("@UNSPSCVersion", project.UNSPSCVersion);

                    //Project MRO Dictionary Version
                    cmd.Parameters.AddWithValue("@MRODictionaryVersion", project.MRODictionaryVersion);

                    //Project Department
                    cmd.Parameters.AddWithValue("@Department", project.Department);

                    //Send Project Activities as xml
                    cmd.Parameters.Add(new SqlParameter("@ProjectActivityDetails", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    cmd.Parameters.AddWithValue("@UserID", project.UserID);
                    #endregion

                    //Calling sp to create project
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "created")
                    {
                        string ProjectCode = arrResult[1];

                        #region Moving all attachments starts

                        #region Customer Input File move Starts
                        if (File.Exists(dirTemp + project.CustomerInputFileName))
                        {
                            string FinalCIFName = project.CustomerCode + "_" + ProjectCode + "_CustomerInputFile" + CustomerInputFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                            FileOperations.MoveFile(dirTemp, project.CustomerInputFileName, dirUploads, FinalCIFName);
                        }
                        #endregion

                        #region Delivery Plan File move Starts
                        if (File.Exists(dirTemp + project.DeliveryPlanFileName))
                        {
                            string FinalDPFName = project.CustomerCode + "_" + ProjectCode + "_DeliveryPlanFile" + DeliveryPlanFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/DeliveryPlanFile/"));
                            FileOperations.MoveFile(dirTemp, project.DeliveryPlanFileName, dirUploads, FinalDPFName);
                        }
                        #endregion

                        #region Scope File move Starts
                        if (File.Exists(dirTemp + project.ScopeFileName))
                        {
                            string FinalScopeFileName = project.CustomerCode + "_" + ProjectCode + "_Scope" + ScopeFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Scope/"));
                            FileOperations.MoveFile(dirTemp, project.ScopeFileName, dirUploads, FinalScopeFileName);
                        }
                        #endregion

                        #region Guideline File move Starts
                        if (File.Exists(dirTemp + project.GuidelineFileName))
                        {
                            string FinalGuidelineFileName = project.CustomerCode + "_" + ProjectCode + "_Guidelines" + GuidelineFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Guidelines/"));
                            FileOperations.MoveFile(dirTemp, project.GuidelineFileName, dirUploads, FinalGuidelineFileName);
                        }
                        #endregion

                        #region Checklist File move Starts
                        if (File.Exists(dirTemp + project.ChecklistFileName))
                        {
                            string FinalChecklistFileName = project.CustomerCode + "_" + ProjectCode + "_Checklist" + ChecklistFileExtension;
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Checklist/"));
                            FileOperations.MoveFile(dirTemp, project.ChecklistFileName, dirUploads, FinalChecklistFileName);
                        }
                        #endregion

                        #endregion

                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, ProjectCode);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read On-Going Projects List
        [HttpPost]
        [Route("ReadOnGoingProjectsList")]
        public IHttpActionResult ReadOnGoingProjectsList(Project projectModel)
        {
            try
            {
                bool canUserChangeProjectStatus = false;
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(projectModel.UserID, "Project List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                if (AccessControl.CanUserAccessPage(projectModel.UserID, "Change Project Status"))
                    canUserChangeProjectStatus = true;

                //Create a list to hold the list of Projects
                List<Project> ProjectsList = new List<Project>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spOnGoingProjectsList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                    cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                    #endregion

                    //Calling sp to get list of projects
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        Project project = new Project();

                        //Assign values to model object
                        project.SlNo = SlNo;
                        project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        project.CustomerCode = sqlReader["CustomerCode"].ToString();
                        project.ProjectCode = sqlReader["ProjectCode"].ToString();
                        project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);
                        project.NoOfActivities = Convert.ToInt32(sqlReader["NoOfActivities"]);
                        project.Scope = sqlReader["Scope"].ToString();
                        project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                        project.ProjectType = sqlReader["ProjectType"].ToString();
                        project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        project.ProductionAllocatedCount = Convert.ToInt64(sqlReader["ProductionAllocatedCount"]);
                        project.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                        project.ProductionPendingCount = Convert.ToInt64(sqlReader["ProductionPendingCount"]);
                        project.QCAllocatedCount = Convert.ToInt64(sqlReader["QCAllocatedCount"]);
                        project.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);
                        project.QCPendingCount = Convert.ToInt64(sqlReader["QCPendingCount"]);
                        project.Status = sqlReader["Status"].ToString();
                        if (project.Status.Trim().ToLower() == "on hold")
                        {
                            project.HoldOnDate = sqlReader["HoldOnDate"] == DBNull.Value ? null : (DateTime?)sqlReader["HoldOnDate"];
                            project.HoldOnReason = sqlReader["HoldOnReason"].ToString();
                        }

                        project.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        project.NoOfResources = Convert.ToInt32(sqlReader["NoOfResources"]);
                        project.canUserChangeProjectStatus = canUserChangeProjectStatus;
                        project.IsProjectSettingsExist = Convert.ToBoolean(sqlReader["IsProjectSettingsExist"]);

                        //Add object to list
                        ProjectsList.Add(project);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectsList);
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

        #region Read Delivered Projects List
        [HttpPost]
        [Route("ReadDeliveredProjectsList")]
        public IHttpActionResult ReadDeliveredProjectsList(Project projectModel)
        {
            try
            {
                bool canUserChangeProjectStatus = false;
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(projectModel.UserID, "Project List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                if (AccessControl.CanUserAccessPage(projectModel.UserID, "Change Project Status"))
                    canUserChangeProjectStatus = true;

                //Create a list to hold the list of Projects
                List<Project> ProjectsList = new List<Project>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spDeliveredProjectsList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                    cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                    #endregion

                    //Calling sp to get list of projects
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        Project project = new Project();

                        //Assign values to model object
                        project.SlNo = SlNo;
                        project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        project.CustomerCode = sqlReader["CustomerCode"].ToString();
                        project.ProjectCode = sqlReader["ProjectCode"].ToString();
                        project.Scope = sqlReader["Scope"].ToString();
                        project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                        project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);
                        project.ProjectType = sqlReader["ProjectType"].ToString();
                        project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        project.Status = sqlReader["Status"].ToString();
                        project.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        project.DeliveredOn = sqlReader["DeliveredOn"] == DBNull.Value ? null : (DateTime?)sqlReader["DeliveredOn"];
                        project.canUserChangeProjectStatus = canUserChangeProjectStatus;

                        //Add object to list
                        ProjectsList.Add(project);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectsList);
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

        #region Read Project By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadProjectById(long id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Project"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                System.Data.Common.DbDataReader sqlReader;

                //Create a Project instance
                Project project = new Project();

                //Create a list to hold Project Activities
                List<ActivityTarget> ActivityTargetList = new List<ActivityTarget>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    #region Project Activities

                    //Initialize Project Activities Command Object
                    SqlCommand cmdProjectActivities = conn.CreateCommand();
                    cmdProjectActivities.CommandText = "spProjectActivities";
                    cmdProjectActivities.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmdProjectActivities.Parameters.AddWithValue("@ProjectID", id);

                    //Call sp to get all project activities
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmdProjectActivities.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create temp activity target object and assign data
                        ActivityTarget activityTarget = new ActivityTarget();
                        activityTarget.Activity = sqlReader["Activity"].ToString();
                        activityTarget.NoOfSKUs = Convert.ToInt32(sqlReader["NoOfSKUs"]);
                        activityTarget.ProductionTarget = Convert.ToInt32(sqlReader["ProductionTarget"]);
                        activityTarget.QCTarget = Convert.ToInt32(sqlReader["QCTarget"]);
                        activityTarget.QATarget = Convert.ToInt32(sqlReader["QATarget"]);
                        activityTarget.AllocatedCount = Convert.ToInt32(sqlReader["AllocatedCount"]);
                        //Add object to list
                        ActivityTargetList.Add(activityTarget);
                    }
                    conn.Close();
                    #endregion

                    #region Project
                    //Initialize Command Object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spViewProject";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", id);

                    //Call sp to get all project details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        //Assign data to project object properties
                        project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        project.CustomerCode = sqlReader["CustomerCode"].ToString();
                        project.ProjectCode = sqlReader["ProjectCode"].ToString();
                        project.LocationCode = sqlReader["LocationCode"].ToString();
                        project.ProjectType = sqlReader["ProjectType"].ToString();
                        project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        project.InputCountType = sqlReader["InputCountType"].ToString();
                        project.TypeOfInput = sqlReader["TypeOfInput"].ToString();
                        project.IsResourceBased = Convert.ToBoolean(sqlReader["IsResourceBased"]);
                        project.PlannedStartDate = Convert.ToDateTime(sqlReader["PlannedStartDate"]);
                        project.Remarks = sqlReader["Remarks"].ToString();
                        project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        project.ReceivedFormat = sqlReader["ReceivedFormat"].ToString();
                        project.OutputFormat = sqlReader["OutputFormat"].ToString();
                        project.DeliveryMode = sqlReader["DeliveryMode"].ToString();
                        project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                        project.CustomerInputFileName = sqlReader["CustomerInputFileName"].ToString();
                        project.DeliveryPlanFileName = sqlReader["DeliveryPlanFileName"].ToString();
                        project.Scope = sqlReader["Scope"].ToString();
                        project.ScopeFileName = sqlReader["ScopeFileName"].ToString();
                        project.Guideline = sqlReader["Guideline"].ToString();
                        project.GuidelineFileName = sqlReader["GuidelineFileName"].ToString();
                        project.Checklist = sqlReader["Checklist"].ToString();
                        project.ChecklistFileName = sqlReader["ChecklistFileName"].ToString();
                        project.EmailDate = sqlReader["EmailDate"] == DBNull.Value ? null : (DateTime?)sqlReader["EmailDate"];
                        project.EmailDescription = sqlReader["EmailDescription"].ToString();
                        project.DeliveredOn = sqlReader["DeliveredOn"] == DBNull.Value ? null : (DateTime?)sqlReader["DeliveredOn"];
                        project.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                        project.Status = sqlReader["ProjectStatus"].ToString();
                        project.IsProjectSettingsExist = Convert.ToBoolean(sqlReader["IsProjectSettingsExist"]);
                        project.UNSPSCVersion = sqlReader["UNSPSCVersion"].ToString();
                        project.MRODictionaryVersion = sqlReader["MRODictionaryNameOrNo"].ToString();
                        project.CreatedByEmployeeName = sqlReader["CreatedByEmployeeName"].ToString();
                        project.Department = sqlReader["Department"].ToString();
                        project.Activities = ActivityTargetList;
                        conn.Close();

                        //return project to the request
                        return Ok(project);
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }
                    #endregion
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

        #region Update Project
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateProject(long id, [FromBody]Project project)
        {
            try
            {
                string CustomerInputFileExtension = string.Empty;
                string DeliveryPlanFileExtension = string.Empty;
                string ScopeFileExtension = string.Empty;
                string GuidelineFileExtension = string.Empty;
                string ChecklistFileExtension = string.Empty;

                //Check Project Code
                if (string.IsNullOrEmpty(project.ProjectCode))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Code");

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Data");

                //Check Project Id
                if (id != project.ProjectID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Id");

                if (!AccessControl.CanUserAccessPage(project.UserID, "Edit Project"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Get all the files extension from temp folder
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                //Customer Input File
                if (File.Exists(dirTemp + project.CustomerInputFileName))
                    CustomerInputFileExtension = Path.GetExtension(dirTemp + project.CustomerInputFileName);
                else if (!string.IsNullOrEmpty(project.CustomerInputFileName))
                    CustomerInputFileExtension = project.CustomerInputFileName.Substring(project.CustomerInputFileName.LastIndexOf('.'), project.CustomerInputFileName.Length - project.CustomerInputFileName.LastIndexOf('.'));


                //Delivery Plan File
                if (File.Exists(dirTemp + project.DeliveryPlanFileName))
                    DeliveryPlanFileExtension = Path.GetExtension(dirTemp + project.DeliveryPlanFileName);
                else if (!string.IsNullOrEmpty(project.DeliveryPlanFileName))
                    DeliveryPlanFileExtension = project.DeliveryPlanFileName.Substring(project.DeliveryPlanFileName.LastIndexOf('.'), project.DeliveryPlanFileName.Length - project.DeliveryPlanFileName.LastIndexOf('.'));

                //Scope File
                if (File.Exists(dirTemp + project.ScopeFileName))
                    ScopeFileExtension = Path.GetExtension(dirTemp + project.ScopeFileName);
                else if (!string.IsNullOrEmpty(project.ScopeFileName))
                    ScopeFileExtension = project.ScopeFileName.Substring(project.ScopeFileName.LastIndexOf('.'), project.ScopeFileName.Length - project.ScopeFileName.LastIndexOf('.'));

                //Guideline File
                if (File.Exists(dirTemp + project.GuidelineFileName))
                    GuidelineFileExtension = Path.GetExtension(dirTemp + project.GuidelineFileName);
                else if (!string.IsNullOrEmpty(project.GuidelineFileName))
                    GuidelineFileExtension = project.GuidelineFileName.Substring(project.GuidelineFileName.LastIndexOf('.'), project.GuidelineFileName.Length - project.GuidelineFileName.LastIndexOf('.'));

                //Checklist File
                if (File.Exists(dirTemp + project.ChecklistFileName))
                    ChecklistFileExtension = Path.GetExtension(dirTemp + project.ChecklistFileName);
                else if (!string.IsNullOrEmpty(project.ChecklistFileName))
                    ChecklistFileExtension = project.ChecklistFileName.Substring(project.ChecklistFileName.LastIndexOf('.'), project.ChecklistFileName.Length - project.ChecklistFileName.LastIndexOf('.'));
                #endregion

                #region Project Activities
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ActivityTarget>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, project.Activities);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spEditProject";

                    #region Adding Stored Procedure Parameters
                    //Project
                    cmd.Parameters.AddWithValue("@CustomerCode", project.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", project.ProjectCode);
                    cmd.Parameters.AddWithValue("@LocationCode", project.LocationCode);
                    cmd.Parameters.AddWithValue("@ProjectType", project.ProjectType);
                    cmd.Parameters.AddWithValue("@InputCount", project.InputCount);
                    cmd.Parameters.AddWithValue("@InputCountType", project.InputCountType);
                    cmd.Parameters.AddWithValue("@TypeOfInput", project.TypeOfInput);
                    cmd.Parameters.AddWithValue("@IsResourceBased", project.IsResourceBased);
                    cmd.Parameters.AddWithValue("@PlannedStartDate", project.PlannedStartDate);
                    cmd.Parameters.AddWithValue("@Remarks", project.Remarks);
                    cmd.Parameters.AddWithValue("@CustomerInputFileExtension", CustomerInputFileExtension);

                    //Project Received Details
                    cmd.Parameters.AddWithValue("@ReceivedDate", project.ReceivedDate);
                    cmd.Parameters.AddWithValue("@ReceivedFormat", project.ReceivedFormat);

                    //Project Delivery Details
                    cmd.Parameters.AddWithValue("@OutputFormat", project.OutputFormat);
                    cmd.Parameters.AddWithValue("@DeliveryMode", project.DeliveryMode);
                    cmd.Parameters.AddWithValue("@PlannedDeliveryDate", project.PlannedDeliveryDate);
                    cmd.Parameters.AddWithValue("@DeliveryPlanFileExtension", DeliveryPlanFileExtension);

                    //Project Scope
                    cmd.Parameters.AddWithValue("@Scope", project.Scope);
                    cmd.Parameters.AddWithValue("@ScopeFileExtension", ScopeFileExtension);

                    //Project Guideline
                    cmd.Parameters.AddWithValue("@Guideline", project.Guideline);
                    cmd.Parameters.AddWithValue("@GuidelineFileExtension", GuidelineFileExtension);

                    //Project Checklist
                    cmd.Parameters.AddWithValue("@Checklist", project.Checklist);
                    cmd.Parameters.AddWithValue("@ChecklistFileExtension", ChecklistFileExtension);

                    //Project Client Details
                    cmd.Parameters.AddWithValue("@EmailDate", project.EmailDate);
                    cmd.Parameters.AddWithValue("@EmailDescription", project.EmailDescription);

                    //Project UNSPSC Version
                    cmd.Parameters.AddWithValue("@UNSPSCVersion", project.UNSPSCVersion);

                    //Project MRO Dictionary Version
                    cmd.Parameters.AddWithValue("@MRODictionaryVersion", project.MRODictionaryVersion);

                    //Project Department
                    cmd.Parameters.AddWithValue("@Department", project.Department);

                    cmd.Parameters.Add(new SqlParameter("@ProjectActivityDetails", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    cmd.Parameters.AddWithValue("@UserID", project.UserID);
                    #endregion

                    //Calling sp to update project
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                    {
                        #region Moving all attachments starts

                        #region Customer Input File
                        DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                        if (File.Exists(dirTemp + project.CustomerInputFileName))
                        {
                            string FinalCIFName = project.CustomerCode + "_" + project.ProjectCode + "_CustomerInputFile" + CustomerInputFileExtension;
                            FileOperations.MoveFile(dirTemp, project.CustomerInputFileName, dirCIFUploads, FinalCIFName);
                        }
                        else if (string.IsNullOrEmpty(project.CustomerInputFileName))
                        {
                            #region Delete Customer Input File
                            foreach (FileInfo file in dirCIFUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(project.CustomerCode.ToUpper() + "_" + project.ProjectCode.ToUpper() + "_CUSTOMERINPUTFILE"))
                                {
                                    file.Delete();
                                    break;
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region Delivery Plan File
                        DirectoryInfo dirDPFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/DeliveryPlanFile/"));
                        if (File.Exists(dirTemp + project.DeliveryPlanFileName))
                        {
                            string FinalDPFName = project.CustomerCode + "_" + project.ProjectCode + "_DeliveryPlanFile" + DeliveryPlanFileExtension;
                            FileOperations.MoveFile(dirTemp, project.DeliveryPlanFileName, dirDPFUploads, FinalDPFName);
                        }
                        else if (string.IsNullOrEmpty(project.DeliveryPlanFileName))
                        {
                            #region Delete Delivery Plan File
                            foreach (FileInfo file in dirDPFUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(project.CustomerCode.ToUpper() + "_" + project.ProjectCode.ToUpper() + "_DELIVERYPLANFILE"))
                                {
                                    file.Delete();
                                    break;
                                }
                            }
                            #endregion
                        }

                        #endregion

                        #region Scope File
                        DirectoryInfo dirScopeFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Scope/"));
                        if (File.Exists(dirTemp + project.ScopeFileName))
                        {
                            string FinalScopeFileName = project.CustomerCode + "_" + project.ProjectCode + "_Scope" + ScopeFileExtension;
                            FileOperations.MoveFile(dirTemp, project.ScopeFileName, dirScopeFileUploads, FinalScopeFileName);
                        }
                        else if (string.IsNullOrEmpty(project.ScopeFileName))
                        {
                            #region Delete Scope File
                            foreach (FileInfo file in dirScopeFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(project.CustomerCode.ToUpper() + "_" + project.ProjectCode.ToUpper() + "_SCOPE"))
                                {
                                    file.Delete();
                                    break;
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region Guideline File
                        DirectoryInfo dirGuidelineFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Guidelines/"));
                        if (File.Exists(dirTemp + project.GuidelineFileName))
                        {
                            string FinalGuidelineFileName = project.CustomerCode + "_" + project.ProjectCode + "_Guidelines" + GuidelineFileExtension;
                            FileOperations.MoveFile(dirTemp, project.GuidelineFileName, dirGuidelineFileUploads, FinalGuidelineFileName);
                        }
                        else if (string.IsNullOrEmpty(project.GuidelineFileName))
                        {
                            #region Delete Guideline File
                            foreach (FileInfo file in dirGuidelineFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(project.CustomerCode.ToUpper() + "_" + project.ProjectCode.ToUpper() + "_GUIDELINES"))
                                {
                                    file.Delete();
                                    break;
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region Checklist File
                        DirectoryInfo dirChecklistFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Checklist/"));
                        if (File.Exists(dirTemp + project.ChecklistFileName))
                        {
                            string FinalChecklistFileName = project.CustomerCode + "_" + project.ProjectCode + "_Checklist" + ChecklistFileExtension;
                            FileOperations.MoveFile(dirTemp, project.ChecklistFileName, dirChecklistFileUploads, FinalChecklistFileName);
                        }
                        else if (string.IsNullOrEmpty(project.ChecklistFileName))
                        {
                            #region Delete Checklist File
                            foreach (FileInfo file in dirChecklistFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(project.CustomerCode.ToUpper() + "_" + project.ProjectCode.ToUpper() + "_CHECKLIST"))
                                {
                                    file.Delete();
                                    break;
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }
        #endregion

        #region Delete Project
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteProject(long id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Project"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spDeleteProject";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    //Calling sp to delete Project
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "deleted")
                    {
                        string CustomerCode = arrResult[1];
                        string ProjectCode = arrResult[2];

                        #region Delete Project Files

                        #region Delete Customer Input File
                        DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                        foreach (FileInfo file in dirCIFUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_CUSTOMERINPUTFILE"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        #region Delete Delivery Plan File
                        DirectoryInfo dirDPFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/DeliveryPlanFile/"));
                        foreach (FileInfo file in dirDPFUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_DELIVERYPLANFILE"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        #region Delete Scope File
                        DirectoryInfo dirScopeFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Scope/"));
                        foreach (FileInfo file in dirScopeFileUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_SCOPE"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        #region Delete Guideline File
                        DirectoryInfo dirGuidelineFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Guidelines/"));
                        foreach (FileInfo file in dirGuidelineFileUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_GUIDELINES"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        #region Delete Checklist File
                        DirectoryInfo dirChecklistFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Checklist/"));
                        foreach (FileInfo file in dirChecklistFileUploads.GetFiles())
                        {
                            if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_CHECKLIST"))
                            {
                                file.Delete();
                                break;
                            }
                        }
                        #endregion

                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region File Upload
        [HttpPost]
        [Route("uploadfile")]
        public HttpResponseMessage UploadFile()
        {
            try
            {
                var httpRequest = HttpContext.Current.Request;
                var postedFile = httpRequest.Files["File"];

                if (postedFile != null)
                {
                    string NewFileName = Guid.NewGuid() + Path.GetExtension(postedFile.FileName);
                    var filePath = HttpContext.Current.Server.MapPath("~/temp/" + NewFileName);
                    postedFile.SaveAs(filePath);

                    return Request.CreateResponse(HttpStatusCode.Created, NewFileName);
                }
                else
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload file.");
                }
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Download file
        [HttpGet]
        [Route("downloadfile")]
        public HttpResponseMessage DownloadFile(string FileName, string FileType = "")
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/temp/") + FileName;

                if (!File.Exists(filePath))
                {
                    if (FileType.ToLower() == "checklist")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Checklist/") + FileName;
                    else if (FileType.ToLower() == "guidelines")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Guidelines/") + FileName;
                    else if (FileType.ToLower() == "scope")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Scope/") + FileName;
                    else if (FileType.ToLower() == "deliveryplanfile")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/DeliveryPlanFile/") + FileName;
                    else if (FileType.ToLower() == "customerinputfile")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/") + FileName;
                    else if (FileType.ToLower() == "profilephoto")
                        filePath = HttpContext.Current.Server.MapPath("~/Uploads/UserImages/") + FileName;
                    else
                    {
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File Type not found");
                    }

                    //Check whether File exists.
                    if (!File.Exists(filePath))
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");
                }

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(filePath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileName));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Delete file
        [HttpPost]
        [Route("deletefile")]
        public HttpResponseMessage DeleteFile(string FileName)
        {
            try
            {
                //Set the File Path.
                string FilePath = HttpContext.Current.Server.MapPath("~/temp/") + FileName;

                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                    return Request.CreateResponse(HttpStatusCode.OK, "deleted");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "");
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Change Project Code
        [HttpPost]
        [Route("ChangeProjectCode/{CustomerCode}/{ProjectCode}/{ChangeToProjectCode}/{UserID}")]
        public HttpResponseMessage ChangeProjectCode(string CustomerCode, string ProjectCode, string ChangeToProjectCode, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Change Project Code"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spChangeProjectCode";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@ChangeToProjectCode", ChangeToProjectCode);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to update project
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "updated")
                    {
                        string ChangedCustomerInputFileName = arrResult[1];
                        string ChangedDeliveryPlanFileName = arrResult[2];
                        string ChangedScopeFileName = arrResult[3];
                        string ChangedGuidelineFileName = arrResult[4];
                        string ChangedChecklistFileName = arrResult[5];

                        #region Renaming Project Files
                        #region Customer Input File
                        if (!string.IsNullOrEmpty(ChangedCustomerInputFileName))
                        {
                            DirectoryInfo dirCIFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                            foreach (FileInfo file in dirCIFUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_CUSTOMERINPUTFILE"))
                                {
                                    file.MoveTo(Path.Combine(dirCIFUploads.FullName, ChangedCustomerInputFileName));
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region Delivery Plan File
                        if (!string.IsNullOrEmpty(ChangedDeliveryPlanFileName))
                        {
                            DirectoryInfo dirDPFUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/DeliveryPlanFile/"));
                            foreach (FileInfo file in dirDPFUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_DELIVERYPLANFILE"))
                                {
                                    file.MoveTo(Path.Combine(dirDPFUploads.FullName, ChangedDeliveryPlanFileName));
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region Scope File
                        if (!string.IsNullOrEmpty(ChangedScopeFileName))
                        {
                            DirectoryInfo dirScopeFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Scope/"));
                            foreach (FileInfo file in dirScopeFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_SCOPE"))
                                {
                                    file.MoveTo(Path.Combine(dirScopeFileUploads.FullName, ChangedScopeFileName));
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region Guideline File
                        if (!string.IsNullOrEmpty(ChangedGuidelineFileName))
                        {
                            DirectoryInfo dirGuidelineFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Guidelines/"));
                            foreach (FileInfo file in dirGuidelineFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_GUIDELINES"))
                                {
                                    file.MoveTo(Path.Combine(dirGuidelineFileUploads.FullName, ChangedGuidelineFileName));
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region Checklist File
                        if (!string.IsNullOrEmpty(ChangedChecklistFileName))
                        {
                            DirectoryInfo dirChecklistFileUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/Checklist/"));
                            foreach (FileInfo file in dirChecklistFileUploads.GetFiles())
                            {
                                if (file.Name.ToUpper().StartsWith(CustomerCode.ToUpper() + "_" + ProjectCode.ToUpper() + "_CHECKLIST"))
                                {
                                    file.MoveTo(Path.Combine(dirChecklistFileUploads.FullName, ChangedChecklistFileName));
                                    break;
                                }
                            }
                        }
                        #endregion
                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK, arrResult[0].Trim().ToLower());
                    }
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Export On-Going Projects List To Excel
        [HttpPost]
        [Route("ExportOnGoingProjectsListToExcel")]
        public HttpResponseMessage ExportOnGoingProjectsListToExcel(Project projectModel)
        {
            try
            {
                if (projectModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("on-going projects list", projectModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = "On-Going Projects List.xlsx";
                    int SlNo = 1;

                    //Create a list to hold the list of Projects
                    List<Project> ProjectsList = new List<Project>();
                    System.Data.Common.DbDataReader sqlReader;

                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Setting up the workbook
                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
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
                        cmd.CommandText = "spOnGoingProjectsList";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        #region Adding Stored Procedure Parameters
                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                        cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                        #endregion

                        //Calling sp to get list of projects
                        conn.Open();
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                //Create model object
                                Project project = new Project();

                                //Assign values to model object
                                project.SlNo = SlNo;
                                project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                                project.CustomerCode = sqlReader["CustomerCode"].ToString();
                                project.ProjectCode = sqlReader["ProjectCode"].ToString();
                                project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);
                                project.NoOfActivities = Convert.ToInt32(sqlReader["NoOfActivities"]);
                                project.Scope = sqlReader["Scope"].ToString();
                                project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                                project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                                project.ProjectType = sqlReader["ProjectType"].ToString();
                                project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                                project.ProductionAllocatedCount = Convert.ToInt64(sqlReader["ProductionAllocatedCount"]);
                                project.ProductionCompletedCount = Convert.ToInt64(sqlReader["ProductionCompletedCount"]);
                                project.ProductionPendingCount = Convert.ToInt64(sqlReader["ProductionPendingCount"]);
                                project.QCAllocatedCount = Convert.ToInt64(sqlReader["QCAllocatedCount"]);
                                project.QCCompletedCount = Convert.ToInt64(sqlReader["QCCompletedCount"]);
                                project.QCPendingCount = Convert.ToInt64(sqlReader["QCPendingCount"]);
                                project.Status = sqlReader["Status"].ToString();
                                if (project.Status.Trim().ToLower() == "on hold")
                                {
                                    project.HoldOnDate = sqlReader["HoldOnDate"] == DBNull.Value ? null : (DateTime?)sqlReader["HoldOnDate"];
                                    project.HoldOnReason = sqlReader["HoldOnReason"].ToString();
                                }

                                project.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                                project.NoOfResources = Convert.ToInt32(sqlReader["NoOfResources"]);

                                //Add object to list
                                ProjectsList.Add(project);
                                SlNo++;
                            }
                            #endregion
                            conn.Close();

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Customer Code");
                            ws.Cells[0, 2].PutValue("Project Code");
                            ws.Cells[0, 3].PutValue("No. of Batches");
                            ws.Cells[0, 4].PutValue("No. of Activities");
                            ws.Cells[0, 5].PutValue("Description");
                            ws.Cells[0, 6].PutValue("Received On");
                            ws.Cells[0, 7].PutValue("Planned Delivery Date");
                            ws.Cells[0, 8].PutValue("Project Type");
                            ws.Cells[0, 9].PutValue("Input Count");
                            ws.Cells[0, 10].PutValue("Prod. Allocated Count");
                            ws.Cells[0, 11].PutValue("Prod. Completed Count");
                            ws.Cells[0, 12].PutValue("Prod. Pending Count");
                            ws.Cells[0, 13].PutValue("QC Allocated Count");
                            ws.Cells[0, 14].PutValue("QC Completed Count");
                            ws.Cells[0, 15].PutValue("QC Pending Count");
                            ws.Cells[0, 16].PutValue("Delivered Count");
                            ws.Cells[0, 17].PutValue("Status");
                            ws.Cells[0, 18].PutValue("No. of Resources");
                            ws.Cells[0, 19].PutValue("On Hold Date");
                            ws.Cells[0, 20].PutValue("On Hold Reason");

                            for (int c = 0; c <= 20; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (Project project in ProjectsList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(row);
                                ws.Cells[row, 1].PutValue(project.CustomerCode);
                                ws.Cells[row, 2].PutValue(project.ProjectCode);
                                ws.Cells[row, 3].PutValue(project.NoOfBatches);
                                ws.Cells[row, 4].PutValue(project.NoOfActivities);
                                ws.Cells[row, 5].PutValue(project.Scope);
                                ws.Cells[row, 6].PutValue(Convert.ToDateTime(project.ReceivedDate).ToString("dd-MMM-yyyy"));
                                if (project.PlannedDeliveryDate != null)
                                    ws.Cells[row, 7].PutValue(Convert.ToDateTime(project.PlannedDeliveryDate).ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 8].PutValue(project.ProjectType);
                                ws.Cells[row, 9].PutValue(project.InputCount);
                                ws.Cells[row, 10].PutValue(project.ProductionAllocatedCount);
                                ws.Cells[row, 11].PutValue(project.ProductionCompletedCount);
                                ws.Cells[row, 12].PutValue(project.ProductionPendingCount);
                                ws.Cells[row, 13].PutValue(project.QCAllocatedCount);
                                ws.Cells[row, 14].PutValue(project.QCCompletedCount);
                                ws.Cells[row, 15].PutValue(project.QCPendingCount);
                                ws.Cells[row, 16].PutValue(project.DeliveredCount);
                                ws.Cells[row, 17].PutValue(project.Status);
                                ws.Cells[row, 18].PutValue(project.NoOfResources);
                                if (project.HoldOnDate != null)
                                {
                                    ws.Cells[row, 19].PutValue(Convert.ToDateTime(project.HoldOnDate).ToString("dd-MMM-yyyy"));
                                    ws.Cells[row, 20].PutValue(project.HoldOnReason);
                                }
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

        #region Export Delivered Projects List To Excel
        [HttpPost]
        [Route("ExportDeliveredProjectsListToExcel")]
        public HttpResponseMessage ExportDeliveredProjectsListToExcel(Project projectModel)
        {
            try
            {
                if (projectModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("delivered Projects List", projectModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = "Delivered Projects List.xlsx";

                    //Create a list to hold the list of Projects
                    List<Project> ProjectsList = new List<Project>();
                    System.Data.Common.DbDataReader sqlReader;

                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Setting up the workbook
                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
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
                        cmd.CommandText = "spDeliveredProjectsList";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        #region Adding Stored Procedure Parameters
                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                        cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                        #endregion

                        //Calling sp to get list of projects
                        conn.Open();
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                //Create model object
                                Project project = new Project();

                                //Assign values to model object
                                project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                                project.CustomerCode = sqlReader["CustomerCode"].ToString();
                                project.ProjectCode = sqlReader["ProjectCode"].ToString();
                                project.Scope = sqlReader["Scope"].ToString();
                                project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                                project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                                project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);
                                project.ProjectType = sqlReader["ProjectType"].ToString();
                                project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                                project.Status = sqlReader["Status"].ToString();
                                project.DeliveredCount = Convert.ToInt32(sqlReader["DeliveredCount"]);
                                project.DeliveredOn = sqlReader["DeliveredOn"] == DBNull.Value ? null : (DateTime?)sqlReader["DeliveredOn"];

                                //Add object to list
                                ProjectsList.Add(project);
                            }
                            #endregion
                            conn.Close();

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Customer Code");
                            ws.Cells[0, 2].PutValue("Project Code");
                            ws.Cells[0, 3].PutValue("Description");
                            ws.Cells[0, 4].PutValue("Received Date");
                            ws.Cells[0, 5].PutValue("Planned Delivery Date");
                            ws.Cells[0, 6].PutValue("No. of Batches");
                            ws.Cells[0, 7].PutValue("Project Type");
                            ws.Cells[0, 8].PutValue("Input Count");
                            ws.Cells[0, 9].PutValue("Status");
                            ws.Cells[0, 10].PutValue("Delivered Count");
                            ws.Cells[0, 11].PutValue("Delivered Date");

                            for (int c = 0; c <= 11; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (Project project in ProjectsList)
                            {
                                #region Writing row data
                                ws.Cells[row, 0].PutValue(row);
                                ws.Cells[row, 1].PutValue(project.CustomerCode);
                                ws.Cells[row, 2].PutValue(project.ProjectCode);
                                ws.Cells[row, 3].PutValue(project.Scope);
                                ws.Cells[row, 4].PutValue(Convert.ToDateTime(project.ReceivedDate).ToString("dd-MMM-yyyy"));
                                if (project.PlannedDeliveryDate != null)
                                    ws.Cells[row, 5].PutValue(Convert.ToDateTime(project.PlannedDeliveryDate).ToString("dd-MMM-yyyy"));
                                ws.Cells[row, 6].PutValue(project.NoOfBatches);
                                ws.Cells[row, 7].PutValue(project.ProjectType);
                                ws.Cells[row, 8].PutValue(project.InputCount);
                                ws.Cells[row, 9].PutValue(project.Status);
                                ws.Cells[row, 10].PutValue(project.DeliveredCount);
                                if (project.DeliveredOn != null)
                                    ws.Cells[row, 11].PutValue(Convert.ToDateTime(project.DeliveredOn).ToString("dd-MMM-yyyy"));
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

        #region Fetch Change Project Status To List
        [HttpGet]
        [Route("FetchChangeProjectStatusToList/{CustomerCode}/{ProjectCode}")]
        public IHttpActionResult FetchChangeProjectStatusToList(string CustomerCode, string ProjectCode)
        {
            try
            {
                //Create a list to hold the list of Project's Status
                List<string> projectStatusList = new List<string>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spChangeProjectStatusToList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);

                    //Calling sp to get list of projects
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        projectStatusList.Add(sqlReader["Status"].ToString());
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectStatusList);
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

        #region Change Project Status
        [HttpPut]
        [Route("ChangeProjectStatus")]
        public HttpResponseMessage ChangeProjectStatus([FromBody]ProjectStatusDetailsModel projectStatusDetailsModel)
        {
            try
            {
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Status Details data");

                if (!AccessControl.CanUserAccessPage(projectStatusDetailsModel.UserID, "Change Project Status"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUpdateProjectStatusDetails";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", projectStatusDetailsModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectStatusDetailsModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@ChangeStatusTo", projectStatusDetailsModel.ChangeStatusTo);
                    cmd.Parameters.AddWithValue("@DeliveredDate", projectStatusDetailsModel.DeliveredDate);
                    cmd.Parameters.AddWithValue("@DeliveredCount", projectStatusDetailsModel.DeliveredCount);
                    cmd.Parameters.AddWithValue("@OnHoldReason", projectStatusDetailsModel.OnHoldReason);
                    cmd.Parameters.AddWithValue("@OnHoldDate", projectStatusDetailsModel.OnHoldDate);
                    cmd.Parameters.AddWithValue("@UserID", projectStatusDetailsModel.UserID);
                    #endregion

                    //Calling sp to change project status
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                    {
                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Revert Project Status
        [HttpPut]
        [Route("RevertProjectStatus/{CustomerCode}/{ProjectCode}/{UserID}")]
        public HttpResponseMessage RevertProjectStatus(string CustomerCode, string ProjectCode, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Change Project Status"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spRevertProjectStatus";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to revert project status
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                    {
                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Project Activities with Hours Worked
        [HttpGet]
        [Route("ReadProjectActivitiesWithHoursWorked/{ProjectID}")]
        public IHttpActionResult ReadProjectActivitiesWithHoursWorked(long ProjectID)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the list of Project Activities With Hours Worked
                List<ProjectActivitiesWithHoursWorked> projectActivitiesWithHoursWorkedList = new List<ProjectActivitiesWithHoursWorked>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectActivitiesWithHoursWorked";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", ProjectID);

                    //Calling sp to get list of Project Activities With Hours Worked
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        ProjectActivitiesWithHoursWorked projectActivitiesWithHoursWorked = new ProjectActivitiesWithHoursWorked();

                        //Assign values to model object
                        projectActivitiesWithHoursWorked.SlNo = SlNo;
                        projectActivitiesWithHoursWorked.ProjectActivityID = Convert.ToInt32(sqlReader["ProjectActivityID"]);
                        projectActivitiesWithHoursWorked.ProjectActivity = sqlReader["ProjectActivity"].ToString();
                        projectActivitiesWithHoursWorked.StartDate = sqlReader["StartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["StartDate"];
                        projectActivitiesWithHoursWorked.EndDate = sqlReader["EndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["EndDate"];
                        projectActivitiesWithHoursWorked.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        projectActivitiesWithHoursWorked.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        projectActivitiesWithHoursWorked.HoursWorked = Convert.ToInt32(sqlReader["HoursWorked"]);

                        //Add object to list
                        projectActivitiesWithHoursWorkedList.Add(projectActivitiesWithHoursWorked);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectActivitiesWithHoursWorkedList);
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

        #region Project Activity Resources with Hours Worked
        [HttpGet]
        [Route("ReadProjectActivityResourcesWithHoursWorked/{ProjectID}/{ProjectActivityID}")]
        public IHttpActionResult ReadProjectActivityResourcesWithHoursWorked(long ProjectID, int ProjectActivityID)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the list of Project Activity Resources with Hours Worked
                List<ProjectActivityResourcesWithHoursWorked> projectActivityResourcesWithHoursWorkedList = new List<ProjectActivityResourcesWithHoursWorked>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectActivityResourcesWithHoursWorked";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectID", ProjectID);
                    cmd.Parameters.AddWithValue("@ProjectActivityID", ProjectActivityID);

                    //Calling sp to get list of Project Activity Resources with Hours Worked
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        ProjectActivityResourcesWithHoursWorked projectActivityResourcesWithHoursWorked = new ProjectActivityResourcesWithHoursWorked();

                        //Assign values to model object
                        projectActivityResourcesWithHoursWorked.SlNo = SlNo;
                        projectActivityResourcesWithHoursWorked.UserID = sqlReader["UserID"].ToString();
                        projectActivityResourcesWithHoursWorked.Username = sqlReader["UserName"].ToString();
                        projectActivityResourcesWithHoursWorked.StartDate = sqlReader["StartDate"] == DBNull.Value ? null : (DateTime?)sqlReader["StartDate"];
                        projectActivityResourcesWithHoursWorked.EndDate = sqlReader["EndDate"] == DBNull.Value ? null : (DateTime?)sqlReader["EndDate"];
                        projectActivityResourcesWithHoursWorked.ProductionCompletedCount = Convert.ToInt32(sqlReader["ProductionCompletedCount"]);
                        projectActivityResourcesWithHoursWorked.QCCompletedCount = Convert.ToInt32(sqlReader["QCCompletedCount"]);
                        projectActivityResourcesWithHoursWorked.HoursWorked = Convert.ToInt32(sqlReader["HoursWorked"]);

                        //Add object to list
                        projectActivityResourcesWithHoursWorkedList.Add(projectActivityResourcesWithHoursWorked);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectActivityResourcesWithHoursWorkedList);
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

        #region Read Not Started Projects List
        [HttpPost]
        [Route("ReadNotStartedProjectsList")]
        public IHttpActionResult ReadNotStartedProjectsList(Project projectModel)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(projectModel.UserID, "Project List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of not started Projects List
                List<Project> NotStartedProjectsList = new List<Project>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spNotStartedProjectsList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                    cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                    #endregion

                    //Calling sp to get list of projects
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        Project project = new Project();

                        //Assign values to model object
                        project.SlNo = SlNo;
                        project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        project.CustomerCode = sqlReader["CustomerCode"].ToString();
                        project.ProjectCode = sqlReader["ProjectCode"].ToString();
                        project.Scope = sqlReader["Scope"].ToString();
                        project.ProjectType = sqlReader["ProjectType"].ToString();
                        project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                        project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                        project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                        project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);

                        //Add object to list
                        NotStartedProjectsList.Add(project);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(NotStartedProjectsList);
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

        #region Export Not Started Projects List To Excel
        [HttpPost]
        [Route("ExportNotStartedProjectsListToExcel")]
        public HttpResponseMessage ExportNotStartedProjectsListToExcel(Project projectModel)
        {
            try
            {
                if (projectModel.UserID.ToLower() != "vic")
                {
                    _BLLAccessControl.SendEmailToManagementAboutExportOfData("not started Projects List", projectModel.UserID);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "This function is disabled on server");
                }
                else
                {
                    string FileName = "Not Started Projects List.xlsx";

                    //Create a list to hold the list of Not Started Projects
                    List<Project> NotStartedProjectsList = new List<Project>();
                    System.Data.Common.DbDataReader sqlReader;

                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Setting up the workbook
                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    //wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
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
                        cmd.CommandText = "spNotStartedProjectsList";
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        #region Adding Stored Procedure Parameters
                        //Add parameters with values
                        cmd.Parameters.AddWithValue("@CustomerCode", projectModel.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", projectModel.ProjectCode);
                        cmd.Parameters.AddWithValue("@ProjectType", projectModel.ProjectType);
                        cmd.Parameters.AddWithValue("@FromDate", projectModel.FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", projectModel.ToDate);
                        #endregion

                        //Calling sp to get list of not started projects
                        conn.Open();
                        sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                        if (sqlReader.HasRows)
                        {
                            #region Adding data to List
                            while (sqlReader.Read())
                            {
                                //Create model object
                                Project project = new Project();

                                //Assign values to model object
                                project.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                                project.CustomerCode = sqlReader["CustomerCode"].ToString();
                                project.ProjectCode = sqlReader["ProjectCode"].ToString();
                                project.Scope = sqlReader["Scope"].ToString();
                                project.ProjectType = sqlReader["ProjectType"].ToString();
                                project.ReceivedDate = Convert.ToDateTime(sqlReader["ReceivedDate"]);
                                project.InputCount = Convert.ToInt32(sqlReader["InputCount"]);
                                project.PlannedDeliveryDate = sqlReader["PlannedDeliveryDate"] == DBNull.Value ? null : (DateTime?)sqlReader["PlannedDeliveryDate"];
                                project.NoOfBatches = Convert.ToInt32(sqlReader["NoOfBatches"]);

                                //Add object to list
                                NotStartedProjectsList.Add(project);
                            }
                            conn.Close();
                            #endregion

                            #region Writing column headings and setting style
                            ws.Cells[0, 0].PutValue("S.No.");
                            ws.Cells[0, 1].PutValue("Customer Code");
                            ws.Cells[0, 2].PutValue("Project Code");
                            ws.Cells[0, 3].PutValue("Description");
                            ws.Cells[0, 4].PutValue("Project Type");
                            ws.Cells[0, 5].PutValue("Received Date");
                            ws.Cells[0, 6].PutValue("Input Count");
                            ws.Cells[0, 7].PutValue("Planned Delivery Date");
                            ws.Cells[0, 8].PutValue("No. of Batches");

                            for (int c = 0; c <= 8; c++)
                                ws.Cells[0, c].SetStyle(styleHeader);
                            #endregion

                            #region Writing row data
                            foreach (Project project in NotStartedProjectsList)
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
                                ws.Cells[row, 8].PutValue(project.NoOfBatches);
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
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Update Project Settings
        [HttpPost]
        [Route("UpdateProjectSettings")]
        public HttpResponseMessage UpdateProjectSettings([FromBody] ProjectSettings projectSettings)
        {
            try
            {
                //TODO: Access Control - Temporary Need to implement for Project Settings page
                if (!AccessControl.CanUserAccessPage(projectSettings.UserID, "Create Project"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Project Settings Special Characters
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<ProjectSettingsSpecialCharacters>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, projectSettings.SpecialCharacters);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectSettings";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", projectSettings.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", projectSettings.ProjectCode);
                    cmd.Parameters.AddWithValue("@AdditionalInfoPrefix", projectSettings.AdditionalInfoPrefix);
                    cmd.Parameters.AddWithValue("@MFRNamePrefix", projectSettings.MFRNamePrefix);
                    cmd.Parameters.AddWithValue("@MFRPNPrefix", projectSettings.MFRPNPrefix);
                    cmd.Parameters.AddWithValue("@VendorNamePrefix", projectSettings.VendorNamePrefix);
                    cmd.Parameters.AddWithValue("@VendorPNPrefix", projectSettings.VendorPNPrefix);
                    cmd.Parameters.AddWithValue("@IsToIncludeAdditionalInfoInShortDesc", projectSettings.IsToIncludeAdditionalInfoInShortDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeAdditionalInfoInLongDesc", projectSettings.IsToIncludeAdditionalInfoInLongDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeMFRNameInShortDesc", projectSettings.IsToIncludeMFRNameInShortDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeMFRNameInLongDesc", projectSettings.IsToIncludeMFRNameInLongDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeMFRPNInShortDesc", projectSettings.IsToIncludeMFRPNInShortDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeMFRPNInLongDesc", projectSettings.IsToIncludeMFRPNInLongDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeVendorNameInShortDesc", projectSettings.IsToIncludeVendorNameInShortDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeVendorNameInLongDesc", projectSettings.IsToIncludeVendorNameInLongDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeVendorPNInShortDesc", projectSettings.IsToIncludeVendorPNInShortDesc);
                    cmd.Parameters.AddWithValue("@IsToIncludeVendorPNInLongDesc", projectSettings.IsToIncludeVendorPNInLongDesc);
                    cmd.Parameters.AddWithValue("@IsToConvertAttributeValueToUppercase", projectSettings.IsToConvertAttributeValueToUppercase);
                    cmd.Parameters.AddWithValue("@MFRNameInputColumnName", projectSettings.MFRNameInputColumnName);
                    cmd.Parameters.AddWithValue("@MFRPNInputColumnName", projectSettings.MFRPNInputColumnName);
                    cmd.Parameters.AddWithValue("@VendorNameInputColumnName", projectSettings.VendorNameInputColumnName);
                    cmd.Parameters.AddWithValue("@VendorPNInputColumnName", projectSettings.VendorPNInputColumnName);
                    cmd.Parameters.AddWithValue("@ShortDescriptionInputColumnName", projectSettings.ShortDescriptionInputColumnName);
                    cmd.Parameters.AddWithValue("@LongDescriptionInputColumnName", projectSettings.LongDescriptionInputColumnName);
                    cmd.Parameters.AddWithValue("@CustomColumnName1", projectSettings.CustomColumnName1);
                    cmd.Parameters.AddWithValue("@CustomColumnName2", projectSettings.CustomColumnName2);
                    cmd.Parameters.AddWithValue("@CustomColumnName3", projectSettings.CustomColumnName3);
                    //Send Project Settings Special Characters as xml
                    cmd.Parameters.Add(new SqlParameter("@ProjectSettingsSpecialCharacters", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });
                    cmd.Parameters.AddWithValue("@UserID", projectSettings.UserID);
                    #endregion

                    //Calling sp to update project settings
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK,Result);
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Project Settings By Project
        [HttpGet]
        [Route("ReadProjectSettingsByProject")]
        public IHttpActionResult ReadProjectSettingsByProject(string CustomerCode, string ProjectCode)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;

                //Create a list to hold Project Settings Special Characters
                List<ProjectSettingsSpecialCharacters> ProjectSettingsSpecialCharactersList = new List<ProjectSettingsSpecialCharacters>();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    #region Project Settings Special Characters
                    //Initialize Project Settings Special Characters Command Object
                    SqlCommand cmdProjectSettingsSpecialCharacters = conn.CreateCommand();
                    cmdProjectSettingsSpecialCharacters.CommandText = "spProjectSettingsSpecialCharacters";
                    cmdProjectSettingsSpecialCharacters.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmdProjectSettingsSpecialCharacters.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmdProjectSettingsSpecialCharacters.Parameters.AddWithValue("@ProjectCode", ProjectCode);

                    //Call sp to get all project Settings Special Characters
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmdProjectSettingsSpecialCharacters.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectSettingsSpecialCharacters projectSettingsSpecialCharacters = new ProjectSettingsSpecialCharacters();
                        projectSettingsSpecialCharacters.Characters = sqlReader["Characters"].ToString();
                        ProjectSettingsSpecialCharactersList.Add(projectSettingsSpecialCharacters);
                    }
                    conn.Close();
                    #endregion

                    //Initialize Command Object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectSettingsByProject";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add Parameters with values
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);

                    //Call sp to get all project settings
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        ProjectSettings projectSettings = new ProjectSettings();

                        projectSettings.CustomerCode = CustomerCode;
                        projectSettings.ProjectCode = ProjectCode;
                        projectSettings.AdditionalInfoPrefix = sqlReader["AdditionalInfoPrefix"].ToString();
                        projectSettings.MFRNamePrefix = sqlReader["MFRNamePrefix"].ToString();
                        projectSettings.MFRPNPrefix = sqlReader["MFRPNPrefix"].ToString();
                        projectSettings.VendorNamePrefix = sqlReader["VendorNamePrefix"].ToString();
                        projectSettings.VendorPNPrefix = sqlReader["VendorPNPrefix"].ToString();
                        projectSettings.IsToIncludeAdditionalInfoInShortDesc = Convert.ToBoolean(sqlReader["IsToIncludeAdditionalInfoInShortDesc"]);
                        projectSettings.IsToIncludeAdditionalInfoInLongDesc = Convert.ToBoolean(sqlReader["IsToIncludeAddionalInfoInLongDesc"]);
                        projectSettings.IsToIncludeMFRNameInShortDesc = Convert.ToBoolean(sqlReader["IsToIncludeMFRNameInShortDesc"]);
                        projectSettings.IsToIncludeMFRNameInLongDesc = Convert.ToBoolean(sqlReader["IsToIncludeMFRNameInLongDesc"]);
                        projectSettings.IsToIncludeMFRPNInShortDesc = Convert.ToBoolean(sqlReader["IsToIncludeMFRPNInShortDesc"]);
                        projectSettings.IsToIncludeMFRPNInLongDesc = Convert.ToBoolean(sqlReader["IsToIncludeMFRPNInLongDesc"]);
                        projectSettings.IsToIncludeVendorNameInShortDesc = Convert.ToBoolean(sqlReader["IsToIncludeVendorNameInShortDesc"]);
                        projectSettings.IsToIncludeVendorNameInLongDesc = Convert.ToBoolean(sqlReader["IsToIncludeVendorNameInLongDesc"]);
                        projectSettings.IsToIncludeVendorPNInShortDesc = Convert.ToBoolean(sqlReader["IsToIncludeVendorPNInShortDesc"]);
                        projectSettings.IsToIncludeVendorPNInLongDesc = Convert.ToBoolean(sqlReader["IsToIncludeVendorPNInLongDesc"]);
                        projectSettings.IsToConvertAttributeValueToUppercase = Convert.ToBoolean(sqlReader["IsToConvertAttributeValueToUppercase"]);
                        projectSettings.MFRNameInputColumnName = sqlReader["MFRNameInputColumnName"].ToString();
                        projectSettings.MFRPNInputColumnName = sqlReader["MFRPNInputColumnName"].ToString();
                        projectSettings.VendorNameInputColumnName = sqlReader["VendorNameInputColumnName"].ToString();
                        projectSettings.VendorPNInputColumnName = sqlReader["VendorPNInputColumnName"].ToString();
                        projectSettings.ShortDescriptionInputColumnName = sqlReader["ShortDescriptionInputColumnName"].ToString();
                        projectSettings.LongDescriptionInputColumnName = sqlReader["LongDescriptionInputColumnName"].ToString();
                        projectSettings.CustomColumnName1 = sqlReader["CustomColumnName1"].ToString();
                        projectSettings.CustomColumnName2 = sqlReader["CustomColumnName2"].ToString();
                        projectSettings.CustomColumnName3 = sqlReader["CustomColumnName3"].ToString();
                        projectSettings.SpecialCharacters = ProjectSettingsSpecialCharactersList;
                        conn.Close();

                        //return project settings to the request
                        return Ok(projectSettings);
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

        #region Read Customer Input File Column Names of Project or Batch
        [HttpGet]
        [Route("ReadCustomerInputFileColumnNamesOfProjectOrBatch/{CustomerCode}/{ProjectCode}/{BatchNo?}")]
        public IHttpActionResult ReadCustomerInputFileColumnNamesOfProjectOrBatch(string CustomerCode, string ProjectCode, string BatchNo="")
        {
            try
            {
                //Swagger fix
                if (BatchNo == "{BatchNo}")
                    BatchNo = "";

                string FileName = string.Empty;
                DirectoryInfo dirCustomerInputFile;
                
                if (string.IsNullOrEmpty(BatchNo))
                    FileName = CustomerCode + '_' + ProjectCode + "_CustomerInputFile.xlsx";
                else
                    FileName = CustomerCode + '_' + ProjectCode + '_' + BatchNo + "_CustomerInputFile.xlsx";

                //Create a list to hold column names List
                List<String> ColumnNamesList = new List<String>();

                #region Check Customer Input File exists
                if (string.IsNullOrEmpty(BatchNo))
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                else
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));
                
                string CustomerInputFilepath = dirCustomerInputFile + FileName;

                if (!File.Exists(CustomerInputFilepath))
                    return Content(HttpStatusCode.BadRequest, "Customer Input File Not Found");
                #endregion

                #region Open the File as Work Book and get first worksheet
                Workbook wbCIF = new Workbook();                       //Customer Input File Work Book
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                //wbCIF.Open(CustomerInputFilepath);
                wbCIF.LoadData(CustomerInputFilepath);
                Worksheet wsCIF = wbCIF.Worksheets[0];
                #endregion

                #region Add first worksheet column names to list
                for (int col = 0; col <= wsCIF.Cells.MaxColumn; col++)
                {
                    if(!string.IsNullOrEmpty(wsCIF.Cells[0, col].StringValue.Trim()))
                        ColumnNamesList.Add(wsCIF.Cells[0,col].StringValue.Trim());
                }
                #endregion

                //return list to the request
                return Ok(ColumnNamesList);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Project Update Details
        #region Save the Project Update Details
        [HttpPost]
        [Route("SaveProjectUpdateDetails")]
        public HttpResponseMessage SaveProjectUpdateDetails([FromBody] ProjectUpdateDetails model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Update Details Data");

                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectUpdateDetails";

                    #region Adding Stored Procedure Parameters
                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectUpdateDetailsID", model.ProjectUpdateDetailsID);
                    cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                    cmd.Parameters.AddWithValue("@Subject", model.Subject);
                    cmd.Parameters.AddWithValue("@Details", model.Details);
                    cmd.Parameters.AddWithValue("@UserUploadedFileName", model.UserUploadedFileName);
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);
                    #endregion

                    //Calling sp to save project update details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    string[] arrResult = Result.Split(',');
                    if (arrResult.Length == 2 && arrResult[0].Trim().ToLower() == "updated")
                    {
                        long ProjectUpdateDetailsID = Convert.ToInt64(arrResult[1]);

                        #region Move Uploaded File
                        if(!string.IsNullOrEmpty(model.UserUploadedFileName))
                        {
                            if (File.Exists(dirTemp + model.UserUploadedTempFileName))
                            {
                                string UploadedNewFileName = ProjectUpdateDetailsID.ToString() + "_" + model.UserUploadedFileName;
                                DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/ProjectUpdateDocs/"));
                                FileOperations.MoveFile(dirTemp, model.UserUploadedTempFileName, dirUploads, UploadedNewFileName);
                            }
                        }
                        #endregion

                        //return response status code
                        return Request.CreateResponse(HttpStatusCode.OK);
                    }
                    else
                    {
                        //return error response status code
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Project Update Details
        [HttpGet]
        [Route("ReadProjectUpdateDetails")]
        public IHttpActionResult ReadProjectUpdateDetails(string CustomerCode, string ProjectCode)
        {
            try
            {
                //Create a list to hold the list of Project Update List
                List<ProjectUpdateDetails> projectUpdateDetailsList = new List<ProjectUpdateDetails>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectUpdateDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);
                    #endregion

                    //Calling sp to get list of project update details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectUpdateDetails projectUpdateDetails= new ProjectUpdateDetails();
                        projectUpdateDetails.ProjectUpdateDetailsID = Convert.ToInt64(sqlReader["ProjectUpdateDetailsID"]);
                        projectUpdateDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectUpdateDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectUpdateDetails.Subject = sqlReader["Subject"].ToString();
                        projectUpdateDetails.Details = sqlReader["Details"].ToString();
                        projectUpdateDetails.UserUploadedFileName = sqlReader["UserUploadedFileName"].ToString();
                        projectUpdateDetailsList.Add(projectUpdateDetails);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(projectUpdateDetailsList);
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

        #region Download Project Update Details uploaded document
        [HttpGet]
        [Route("DownloadProjectUpdateDetailsUploadedDocument")]
        public HttpResponseMessage DownloadProjectUpdateDetailsUploadedDocument(long ProjectUpdateDetailsID, string FileName)
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedFileName = ProjectUpdateDetailsID.ToString() + "_" + FileName;
                string UploadedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/ProjectUpdateDocs/") + UploadedFileName;

                //Check whether File exists.
                if (!File.Exists(UploadedFilePath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found");

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(UploadedFilePath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = UploadedFileName;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(UploadedFileName));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Delete Project Update Details
        [HttpPatch]
        [Route("DeleteProjectUpdateDetails")]
        public HttpResponseMessage DeleteProjectUpdateDetails(long id,string UploadedFileName, string UserID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectUpdateDetails";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ProjectUpdateDetailsID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Project Update Details
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "deleted")
                    {
                        if (!string.IsNullOrEmpty(UploadedFileName))
                        {
                            string UploadedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/GPMT/ProjectUpdateDocs/") + id.ToString() + "_" + UploadedFileName;
                            if (File.Exists(UploadedFilePath))
                                File.Delete(UploadedFilePath);
                        }
                    }

                    return Request.CreateResponse(HttpStatusCode.OK);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
        #endregion
    }
}
