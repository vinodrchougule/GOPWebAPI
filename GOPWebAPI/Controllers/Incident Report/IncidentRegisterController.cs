using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models.Incident_Report_Models;
using GOPWebAPI.Models;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Web;
using Aspose.Cells;
using System.Net.Http.Headers;
using System.IO;

namespace GOPWebAPI.Controllers.Incident_Report
{
    [RoutePrefix("api/IncidentRegister")]
    public class IncidentRegisterController : ApiController
    {
        private BLLIncidentRegister _BLLIncidentRegister;
        public IncidentRegisterController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLIncidentRegister = new BLLIncidentRegister(connectionString);
        }

        #region Register an Incident
        [HttpPost]
        [Route("PostRegisterIncident")]
        public HttpResponseMessage PostRegisterIncident([FromBody] IncidentRegisterModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Register Incident"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objResult = new Result();

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Model Data");

                // Call the service to register an incident
                string strOutput = _BLLIncidentRegister.RegisterIncident(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objResult.Msg = strOutput;
                    objResult.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objResult);
                }
                else
                {
                    objResult.Msg = strOutput;
                    objResult.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objResult);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Read All Incidents
        [HttpGet]
        [Route("ReadAllIncidents")]
        public HttpResponseMessage ReadAllIncidents(string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Report"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DataTable dtDetails = new DataTable();
                List<IncidentRegisterModel> lstIncidents = new List<IncidentRegisterModel>();
                dtDetails = _BLLIncidentRegister.ReadAllIncidents();

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentRegisterModel objIncidentRegisterModel = new IncidentRegisterModel();
                        objIncidentRegisterModel.IncidentRegisterID = dtDetails.Rows[i]["IncidentRegisterID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentRegisterID"]) : 0;
                        objIncidentRegisterModel.IncidentNo = dtDetails.Rows[i]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[i]["IncidentNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        if(dtDetails.Rows[i]["IncidentDate"]!=DBNull.Value)
                            objIncidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[i]["IncidentDate"]);
                        objIncidentRegisterModel.IncidentTime = dtDetails.Rows[i]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[i]["IncidentTime"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentStatus = dtDetails.Rows[i]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[i]["IncidentStatus"].ToString().Trim() : "";
                        objIncidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentDescription = dtDetails.Rows[i]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[i]["IncidentDescription"].ToString().Trim() : "";
                        objIncidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[i]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[i]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                        objIncidentRegisterModel.ContactNo = dtDetails.Rows[i]["ContactNo"] != DBNull.Value ? dtDetails.Rows[i]["ContactNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.EmailID = dtDetails.Rows[i]["EmailID"] != DBNull.Value ? dtDetails.Rows[i]["EmailID"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentLocation = dtDetails.Rows[i]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[i]["IncidentLocation"].ToString().Trim() : "";
                        objIncidentRegisterModel.InformationAffected = dtDetails.Rows[i]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[i]["InformationAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.EquipmentAffected = dtDetails.Rows[i]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[i]["EquipmentAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[i]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[i]["ImpactOnBusiness"].ToString().Trim() : "";
                        objIncidentRegisterModel.Priority = dtDetails.Rows[i]["Priority"] != DBNull.Value ? dtDetails.Rows[i]["Priority"].ToString().Trim() : "";
                        objIncidentRegisterModel.AssetIDs = dtDetails.Rows[i]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[i]["AssetIDs"].ToString().Trim() : "";
                        objIncidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[i]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["NoOfPeopleAffected"]) : 0;
                        objIncidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsConfirmed"]);
                        objIncidentRegisterModel.RootCause = dtDetails.Rows[i]["RootCause"] != DBNull.Value ? dtDetails.Rows[i]["RootCause"].ToString().Trim() : "";
                        objIncidentRegisterModel.CorrectiveAction = dtDetails.Rows[i]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[i]["CorrectiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.PreventiveAction = dtDetails.Rows[i]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[i]["PreventiveAction"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["ActionCompletedByUserID"] != DBNull.Value)
                        {
                            objIncidentRegisterModel.ActionCompletedByUserID = Convert.ToInt32(dtDetails.Rows[i]["ActionCompletedByUserID"]);
                            objIncidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[i]["ActionCompletedByUserName"].ToString().Trim();
                        }
                        if (dtDetails.Rows[i]["ActionCompletedOn"] != DBNull.Value)
                            objIncidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[i]["ActionCompletedOn"]);
                        objIncidentRegisterModel.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        objIncidentRegisterModel.IsActionConfirmed = dtDetails.Rows[i]["IsActionConfirmed"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsActionConfirmed"]) : false;
                        lstIncidents.Add(objIncidentRegisterModel);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("Incidents List", new JArray(from p in lstIncidents
                                                               select new JObject(
                                                                new JProperty("IncidentRegisterID", p.IncidentRegisterID),
                                                                new JProperty("IncidentNo", p.IncidentNo),
                                                                new JProperty("IncidentType", p.IncidentType),
                                                                new JProperty("IncidentDate", p.IncidentDate),
                                                                new JProperty("IncidentTime", p.IncidentTime),
                                                                new JProperty("IncidentStatus", p.IncidentStatus),
                                                                new JProperty("IncidentDescription", p.IncidentDescription),
                                                                new JProperty("DepartmentResolvingIncident", p.DepartmentResolvingIncident),
                                                                new JProperty("NameOfPersonReportingIncident", p.NameOfPersonReportingIncident),
                                                                new JProperty("ContactNo", p.ContactNo),
                                                                new JProperty("EmailID", p.EmailID),
                                                                new JProperty("IncidentLocation", p.IncidentLocation),
                                                                new JProperty("InformationAffected", p.InformationAffected),
                                                                new JProperty("EquipmentAffected", p.EquipmentAffected),
                                                                new JProperty("ImpactOnBusiness", p.ImpactOnBusiness),
                                                                new JProperty("Priority", p.Priority),
                                                                new JProperty("AssetIDs", p.AssetIDs),
                                                                new JProperty("NoOfPeopleAffected", p.NoOfPeopleAffected),
                                                                new JProperty("IsConfirmed", p.IsConfirmed),
                                                                new JProperty("RootCause", p.RootCause),
                                                                new JProperty("CorrectiveAction", p.CorrectiveAction),
                                                                new JProperty("PreventiveAction", p.PreventiveAction),
                                                                new JProperty("ActionCompletedByUserID", p.ActionCompletedByUserID),
                                                                new JProperty("ActionCompletedByUserName", p.ActionCompletedByUserName),
                                                                new JProperty("ActionCompletedOn", p.ActionCompletedOn),
                                                                new JProperty("Remarks", p.Remarks),
                                                                new JProperty("IsActionConfirmed", p.IsActionConfirmed)
                                                                ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incidents Unique Search Values by Search Field
        [HttpGet]
        [Route("ReadIncidentsUniqueSearchValuesBySearchField")]
        public HttpResponseMessage ReadIncidentsUniqueSearchValuesBySearchField(string SearchField)
        {
            try
            {
                DataTable dtDetails = new DataTable();
                List<SearchIncidentUniqueFieldValues> lstIncidentsUniqueFieldValues = new List<SearchIncidentUniqueFieldValues>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsUniqueSearchValuesBySearchField(SearchField);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        SearchIncidentUniqueFieldValues searchIncidentUniqueFieldValues = new SearchIncidentUniqueFieldValues();
                        searchIncidentUniqueFieldValues.SearchValue = dtDetails.Rows[i]["SearchResult"] != DBNull.Value ? dtDetails.Rows[i]["SearchResult"].ToString().Trim() : "";
                        lstIncidentsUniqueFieldValues.Add(searchIncidentUniqueFieldValues);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("IncidentsSearchValues", new JArray(from p in lstIncidentsUniqueFieldValues
                                                                        select new JObject(
                                                                             new JProperty("SearchValue", p.SearchValue)))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incident By Id
        [HttpGet]
        [Route("ReadIncidentById")]
        public HttpResponseMessage ReadIncidentById(long incidentRegisterID)
        {
            try
            {
                DataTable dtDetails = new DataTable();
                IncidentRegisterModel incidentRegisterModel = new IncidentRegisterModel();
                dtDetails = _BLLIncidentRegister.ReadIncidentById(incidentRegisterID);
                string DepartmentsAffectedCSV = string.Empty;

                DataTable dtDepartmentsAffected = new DataTable();
                List<DepartmentsAffected> lstDepartmentsAffected = new List<DepartmentsAffected>();

                if (dtDetails.Rows.Count > 0)
                {
                    dtDepartmentsAffected = _BLLIncidentRegister.ReadDepartmentsAffectedById(incidentRegisterID);
                    DepartmentsAffectedCSV = string.Empty;
                    for (int j = 0; j < dtDepartmentsAffected.Rows.Count; j++)
                    {
                        DepartmentsAffected departmentsAffected = new DepartmentsAffected();
                        departmentsAffected.DepartmentName = dtDepartmentsAffected.Rows[j]["Name"].ToString();
                        lstDepartmentsAffected.Add(departmentsAffected);
                        DepartmentsAffectedCSV += departmentsAffected.DepartmentName + ",";
                    }

                    if (DepartmentsAffectedCSV.EndsWith(","))
                        DepartmentsAffectedCSV = DepartmentsAffectedCSV.TrimEnd(',');

                    incidentRegisterModel.IncidentRegisterID = Convert.ToInt64(dtDetails.Rows[0]["IncidentRegisterID"]);
                    incidentRegisterModel.IncidentNo = dtDetails.Rows[0]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[0]["IncidentNo"].ToString().Trim() : "";
                    incidentRegisterModel.IncidentTypeID = Convert.ToInt32(dtDetails.Rows[0]["IncidentTypeID"]);
                    incidentRegisterModel.IncidentType = dtDetails.Rows[0]["IncidentType"] != DBNull.Value ? dtDetails.Rows[0]["IncidentType"].ToString().Trim() : "";
                    incidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[0]["IncidentDate"]);
                    incidentRegisterModel.IncidentTime = dtDetails.Rows[0]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[0]["IncidentTime"].ToString() : "";
                    incidentRegisterModel.IncidentStatus = dtDetails.Rows[0]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[0]["IncidentStatus"].ToString().Trim() : "";
                    incidentRegisterModel.DepartmentIDResolvingIncident = Convert.ToInt32(dtDetails.Rows[0]["DepartmentID"]);
                    incidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[0]["Department"] != DBNull.Value ? dtDetails.Rows[0]["Department"].ToString().Trim() : "";
                    incidentRegisterModel.IncidentDescription = dtDetails.Rows[0]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[0]["IncidentDescription"].ToString().Trim() : "";
                    incidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[0]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[0]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                    incidentRegisterModel.ContactNo = dtDetails.Rows[0]["ContactNo"] != DBNull.Value ? dtDetails.Rows[0]["ContactNo"].ToString().Trim() : "";
                    incidentRegisterModel.EmailID = dtDetails.Rows[0]["EmailID"] != DBNull.Value ? dtDetails.Rows[0]["EmailID"].ToString().Trim() : "";
                    incidentRegisterModel.IncidentLocation = dtDetails.Rows[0]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[0]["IncidentLocation"].ToString().Trim() : "";
                    incidentRegisterModel.InformationAffected = dtDetails.Rows[0]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[0]["InformationAffected"].ToString().Trim() : "";
                    incidentRegisterModel.EquipmentAffected = dtDetails.Rows[0]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[0]["EquipmentAffected"].ToString().Trim() : "";
                    incidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[0]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[0]["ImpactOnBusiness"].ToString().Trim() : "";
                    incidentRegisterModel.Priority = dtDetails.Rows[0]["Priority"] != DBNull.Value ? dtDetails.Rows[0]["Priority"].ToString().Trim() : "";
                    incidentRegisterModel.AssetIDs = dtDetails.Rows[0]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[0]["AssetIDs"].ToString().Trim() : "";
                    incidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[0]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[0]["NoOfPeopleAffected"]) : 0;
                    incidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[0]["IsConfirmed"]);
                    incidentRegisterModel.DepartmentsAffectedList = lstDepartmentsAffected;
                    incidentRegisterModel.RootCause = dtDetails.Rows[0]["RootCause"] != DBNull.Value ? dtDetails.Rows[0]["RootCause"].ToString().Trim() : "";
                    incidentRegisterModel.CorrectiveAction = dtDetails.Rows[0]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[0]["CorrectiveAction"].ToString().Trim() : "";
                    incidentRegisterModel.PreventiveAction = dtDetails.Rows[0]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[0]["PreventiveAction"].ToString().Trim() : "";
                    if (dtDetails.Rows[0]["ActionCompletedByUserID"] != DBNull.Value)
                    {
                        incidentRegisterModel.ActionCompletedByUserID = Convert.ToInt32(dtDetails.Rows[0]["ActionCompletedByUserID"]);
                        incidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[0]["ActionCompletedByUserName"].ToString().Trim();
                    }
                    if(dtDetails.Rows[0]["ActionCompletedOn"]!= DBNull.Value)
                        incidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[0]["ActionCompletedOn"]);
                    incidentRegisterModel.Remarks = dtDetails.Rows[0]["Remarks"] != DBNull.Value ? dtDetails.Rows[0]["Remarks"].ToString().Trim() : "";
                    incidentRegisterModel.IsActionConfirmed = dtDetails.Rows[0]["IsActionConfirmed"] !=DBNull.Value?Convert.ToBoolean(dtDetails.Rows[0]["IsActionConfirmed"]):false;
                    incidentRegisterModel.DepartmentsAffectedCSV = DepartmentsAffectedCSV;

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                                     new JProperty("IncidentDetails", new JObject(
                                                                       new JProperty("IncidentRegisterID", incidentRegisterModel.IncidentRegisterID),
                                                                       new JProperty("IncidentTypeID", incidentRegisterModel.IncidentTypeID),
                                                                       new JProperty("IncidentType", incidentRegisterModel.IncidentType),
                                                                       new JProperty("IncidentNo", incidentRegisterModel.IncidentNo),
                                                                       new JProperty("IncidentDate", incidentRegisterModel.IncidentDate),
                                                                       new JProperty("IncidentDescription", incidentRegisterModel.IncidentDescription),
                                                                       new JProperty("IncidentTime", incidentRegisterModel.IncidentTime),
                                                                       new JProperty("IncidentStatus", incidentRegisterModel.IncidentStatus),
                                                                       new JProperty("DepartmentIDResolvingIncident", incidentRegisterModel.DepartmentIDResolvingIncident),
                                                                       new JProperty("DepartmentResolvingIncident", incidentRegisterModel.DepartmentResolvingIncident),
                                                                       new JProperty("NameOfPersonReportingIncident", incidentRegisterModel.NameOfPersonReportingIncident),
                                                                       new JProperty("ContactNo", incidentRegisterModel.ContactNo),
                                                                       new JProperty("EmailID", incidentRegisterModel.EmailID),
                                                                       new JProperty("IncidentLocation", incidentRegisterModel.IncidentLocation),
                                                                       new JProperty("InformationAffected", incidentRegisterModel.InformationAffected),
                                                                       new JProperty("EquipmentAffected", incidentRegisterModel.EquipmentAffected),
                                                                       new JProperty("ImpactOnBusiness", incidentRegisterModel.ImpactOnBusiness),
                                                                       new JProperty("Priority", incidentRegisterModel.Priority),
                                                                       new JProperty("AssetIDs", incidentRegisterModel.AssetIDs),
                                                                       new JProperty("NoOfPeopleAffected", incidentRegisterModel.NoOfPeopleAffected),
                                                                       new JProperty("IsConfirmed", incidentRegisterModel.IsConfirmed),
                                                                       new JProperty("RootCause", incidentRegisterModel.RootCause),
                                                                       new JProperty("CorrectiveAction", incidentRegisterModel.CorrectiveAction),
                                                                       new JProperty("PreventiveAction", incidentRegisterModel.PreventiveAction),
                                                                       new JProperty("ActionCompletedByUserID", incidentRegisterModel.ActionCompletedByUserID),
                                                                       new JProperty("ActionCompletedByUserName", incidentRegisterModel.ActionCompletedByUserName),
                                                                       new JProperty("ActionCompletedOn", incidentRegisterModel.ActionCompletedOn),
                                                                       new JProperty("Remarks", incidentRegisterModel.Remarks),
                                                                       new JProperty("IsActionConfirmed", incidentRegisterModel.IsActionConfirmed),
                                                                       new JProperty("DepartmentsAffectedCSV", incidentRegisterModel.DepartmentsAffectedCSV),
                                                                       new JProperty("DepartmentsAffected", new JArray(
                                                                           from d in lstDepartmentsAffected
                                                                           select new JObject(
                                                                               new JProperty("DepartmentName", d.DepartmentName)))
                                                                       ))));
                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incidents By Search Field and Search Value
        [HttpGet]
        [Route("ReadIncidentsBySearchFieldAndSearchValue")]
        public HttpResponseMessage ReadIncidentsBySearchFieldAndSearchValue(string SearchField, string SearchValue)
        {
            try
            {
                if(string.IsNullOrEmpty(SearchField) || string.IsNullOrEmpty(SearchValue))
                {
                    Result objResult = new Result();
                    objResult.Msg = "Please select Search Field and Search Value";
                    objResult.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.BadRequest, objResult);
                }

                DataTable dtDetails = new DataTable();
                List<IncidentRegisterModel> lstIncidents = new List<IncidentRegisterModel>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsBySearchFieldAndSearchValue(SearchField, SearchValue);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentRegisterModel objIncidentRegisterModel = new IncidentRegisterModel();
                        objIncidentRegisterModel.IncidentRegisterID = dtDetails.Rows[i]["IncidentRegisterID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentRegisterID"]) : 0;
                        objIncidentRegisterModel.IncidentNo = dtDetails.Rows[i]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[i]["IncidentNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["IncidentDate"] != DBNull.Value)
                            objIncidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[i]["IncidentDate"]);
                        objIncidentRegisterModel.IncidentTime = dtDetails.Rows[i]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[i]["IncidentTime"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentStatus = dtDetails.Rows[i]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[i]["IncidentStatus"].ToString().Trim() : "";
                        objIncidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentDescription = dtDetails.Rows[i]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[i]["IncidentDescription"].ToString().Trim() : "";
                        objIncidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[i]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[i]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                        objIncidentRegisterModel.ContactNo = dtDetails.Rows[i]["ContactNo"] != DBNull.Value ? dtDetails.Rows[i]["ContactNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.EmailID = dtDetails.Rows[i]["EmailID"] != DBNull.Value ? dtDetails.Rows[i]["EmailID"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentLocation = dtDetails.Rows[i]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[i]["IncidentLocation"].ToString().Trim() : "";
                        objIncidentRegisterModel.InformationAffected = dtDetails.Rows[i]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[i]["InformationAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.EquipmentAffected = dtDetails.Rows[i]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[i]["EquipmentAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[i]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[i]["ImpactOnBusiness"].ToString().Trim() : "";
                        objIncidentRegisterModel.Priority = dtDetails.Rows[i]["Priority"] != DBNull.Value ? dtDetails.Rows[i]["Priority"].ToString().Trim() : "";
                        objIncidentRegisterModel.AssetIDs = dtDetails.Rows[i]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[i]["AssetIDs"].ToString().Trim() : "";
                        objIncidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[i]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["NoOfPeopleAffected"]) : 0;
                        objIncidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsConfirmed"]);
                        objIncidentRegisterModel.RootCause = dtDetails.Rows[i]["RootCause"] != DBNull.Value ? dtDetails.Rows[i]["RootCause"].ToString().Trim() : "";
                        objIncidentRegisterModel.CorrectiveAction = dtDetails.Rows[i]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[i]["CorrectiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.PreventiveAction = dtDetails.Rows[i]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[i]["PreventiveAction"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["ActionCompletedByUserID"] != DBNull.Value)
                        {
                            objIncidentRegisterModel.ActionCompletedByUserID = Convert.ToInt32(dtDetails.Rows[i]["ActionCompletedByUserID"]);
                            objIncidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[i]["ActionCompletedByUserName"].ToString().Trim();
                        }
                        if (dtDetails.Rows[i]["ActionCompletedOn"] != DBNull.Value)
                            objIncidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[i]["ActionCompletedOn"]);
                        objIncidentRegisterModel.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        objIncidentRegisterModel.IsActionConfirmed = dtDetails.Rows[i]["IsActionConfirmed"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsActionConfirmed"]) : false;
                        lstIncidents.Add(objIncidentRegisterModel);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("IncidentsList", new JArray(from p in lstIncidents
                                                               select new JObject(
                                                                new JProperty("IncidentRegisterID", p.IncidentRegisterID),
                                                                new JProperty("IncidentNo", p.IncidentNo),
                                                                new JProperty("IncidentType", p.IncidentType),
                                                                new JProperty("IncidentDate", p.IncidentDate),
                                                                new JProperty("IncidentTime", p.IncidentTime),
                                                                new JProperty("IncidentStatus", p.IncidentStatus),
                                                                new JProperty("IncidentDescription", p.IncidentDescription),
                                                                new JProperty("DepartmentResolvingIncident", p.DepartmentResolvingIncident),
                                                                new JProperty("NameOfPersonReportingIncident", p.NameOfPersonReportingIncident),
                                                                new JProperty("ContactNo", p.ContactNo),
                                                                new JProperty("EmailID", p.EmailID),
                                                                new JProperty("IncidentLocation", p.IncidentLocation),
                                                                new JProperty("InformationAffected", p.InformationAffected),
                                                                new JProperty("EquipmentAffected", p.EquipmentAffected),
                                                                new JProperty("ImpactOnBusiness", p.ImpactOnBusiness),
                                                                new JProperty("Priority", p.Priority),
                                                                new JProperty("AssetIDs", p.AssetIDs),
                                                                new JProperty("NoOfPeopleAffected", p.NoOfPeopleAffected),
                                                                new JProperty("IsConfirmed", p.IsConfirmed),
                                                                new JProperty("RootCause", p.RootCause),
                                                                new JProperty("CorrectiveAction", p.CorrectiveAction),
                                                                new JProperty("PreventiveAction", p.PreventiveAction),
                                                                new JProperty("ActionCompletedByUserID", p.ActionCompletedByUserID),
                                                                new JProperty("ActionCompletedOn", p.ActionCompletedOn),
                                                                new JProperty("Remarks", p.Remarks),
                                                                new JProperty("IsActionConfirmed", p.IsActionConfirmed)
                                                                ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incidents Count Summary By Year
        [HttpGet]
        [Route("ReadIncidentsCountSummaryByYear")]
        public HttpResponseMessage ReadIncidentsCountSummaryByYear(string UserID, string YearOfIncident = null)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Report Dashboard"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DataTable dtDetails = new DataTable();
                List<IncidentCountSummary> lstIncidents = new List<IncidentCountSummary>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsCountSummaryByYear(YearOfIncident);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentCountSummary incidentCountSummary= new IncidentCountSummary();
                        incidentCountSummary.Department = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        incidentCountSummary.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        incidentCountSummary.PendingCount = dtDetails.Rows[i]["PendingCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["PendingCount"]) : 0;
                        incidentCountSummary.InProgressCount = dtDetails.Rows[i]["InProgressCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["InProgressCount"]) : 0;
                        incidentCountSummary.CompletedCount = dtDetails.Rows[i]["CompletedCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["CompletedCount"]) : 0;
                        incidentCountSummary.TotalCount = dtDetails.Rows[i]["TotalCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["TotalCount"]) : 0;
                        lstIncidents.Add(incidentCountSummary);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("IncidentsCountSummary", new JArray(from p in lstIncidents
                                                               select new JObject(
                                                                new JProperty("Department", p.Department),
                                                                new JProperty("IncidentType", p.IncidentType),
                                                                new JProperty("PendingCount", p.PendingCount),
                                                                new JProperty("InProgressCount", p.InProgressCount),
                                                                new JProperty("CompletedCount", p.CompletedCount),
                                                                new JProperty("TotalCount", p.TotalCount)
                                                            ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incidents By Year and Status
        [HttpGet]
        [Route("ReadIncidentsByIncidentYearAndStatus")]
        public HttpResponseMessage ReadIncidentsByIncidentYearAndStatus(string Department, string IncidentType, string Status, string YearOfIncident = null)
        {
            try
            {
                if (YearOfIncident != null && YearOfIncident.Trim().ToLower() == "all")
                    YearOfIncident = null;

                DataTable dtDetails = new DataTable();
                List<IncidentRegisterModel> lstIncidents = new List<IncidentRegisterModel>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsByIncidentYearAndStatus(Department, IncidentType, YearOfIncident, Status);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentRegisterModel objIncidentRegisterModel = new IncidentRegisterModel();
                        objIncidentRegisterModel.IncidentRegisterID = dtDetails.Rows[i]["IncidentRegisterID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentRegisterID"]) : 0;
                        objIncidentRegisterModel.IncidentNo = dtDetails.Rows[i]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[i]["IncidentNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["IncidentDate"] != DBNull.Value)
                            objIncidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[i]["IncidentDate"]);
                        objIncidentRegisterModel.IncidentTime = dtDetails.Rows[i]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[i]["IncidentTime"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentStatus = dtDetails.Rows[i]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[i]["IncidentStatus"].ToString().Trim() : "";
                        objIncidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentDescription = dtDetails.Rows[i]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[i]["IncidentDescription"].ToString().Trim() : "";
                        objIncidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[i]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[i]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                        objIncidentRegisterModel.ContactNo = dtDetails.Rows[i]["ContactNo"] != DBNull.Value ? dtDetails.Rows[i]["ContactNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.EmailID = dtDetails.Rows[i]["EmailID"] != DBNull.Value ? dtDetails.Rows[i]["EmailID"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentLocation = dtDetails.Rows[i]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[i]["IncidentLocation"].ToString().Trim() : "";
                        objIncidentRegisterModel.InformationAffected = dtDetails.Rows[i]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[i]["InformationAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.EquipmentAffected = dtDetails.Rows[i]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[i]["EquipmentAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[i]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[i]["ImpactOnBusiness"].ToString().Trim() : "";
                        objIncidentRegisterModel.Priority = dtDetails.Rows[i]["Priority"] != DBNull.Value ? dtDetails.Rows[i]["Priority"].ToString().Trim() : "";
                        objIncidentRegisterModel.AssetIDs = dtDetails.Rows[i]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[i]["AssetIDs"].ToString().Trim() : "";
                        objIncidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[i]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["NoOfPeopleAffected"]) : 0;
                        if (dtDetails.Rows[i]["IsConfirmed"] != DBNull.Value)
                            objIncidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsConfirmed"]);
                        objIncidentRegisterModel.RootCause = dtDetails.Rows[i]["RootCause"] != DBNull.Value ? dtDetails.Rows[i]["RootCause"].ToString().Trim() : "";
                        objIncidentRegisterModel.CorrectiveAction = dtDetails.Rows[i]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[i]["CorrectiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.PreventiveAction = dtDetails.Rows[i]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[i]["PreventiveAction"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["ActionCompletedByUserID"] != DBNull.Value)
                        {
                            objIncidentRegisterModel.ActionCompletedByUserID = Convert.ToInt32(dtDetails.Rows[i]["ActionCompletedByUserID"]);
                            objIncidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[i]["ActionCompletedByUserName"].ToString().Trim();
                        }
                        if (dtDetails.Rows[i]["ActionCompletedOn"] != DBNull.Value)
                            objIncidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[i]["ActionCompletedOn"]);
                        objIncidentRegisterModel.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        if(dtDetails.Rows[i]["IsActionConfirmed"]!=DBNull.Value)
                            objIncidentRegisterModel.IsActionConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsActionConfirmed"]);
                        lstIncidents.Add(objIncidentRegisterModel);
                    }

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                    new JProperty("RecordCount", dtDetails.Rows.Count),
                    new JProperty("IncidentsListByYearAndStatus", new JArray(from p in lstIncidents
                                                               select new JObject(
                                                                new JProperty("IncidentRegisterID", p.IncidentRegisterID),
                                                                new JProperty("IncidentNo", p.IncidentNo),
                                                                new JProperty("IncidentType", p.IncidentType),
                                                                new JProperty("IncidentDate", p.IncidentDate),
                                                                new JProperty("IncidentTime", p.IncidentTime),
                                                                new JProperty("IncidentStatus", p.IncidentStatus),
                                                                new JProperty("IncidentDescription", p.IncidentDescription),
                                                                new JProperty("DepartmentResolvingIncident", p.DepartmentResolvingIncident),
                                                                new JProperty("NameOfPersonReportingIncident", p.NameOfPersonReportingIncident),
                                                                new JProperty("ContactNo", p.ContactNo),
                                                                new JProperty("EmailID", p.EmailID),
                                                                new JProperty("IncidentLocation", p.IncidentLocation),
                                                                new JProperty("InformationAffected", p.InformationAffected),
                                                                new JProperty("EquipmentAffected", p.EquipmentAffected),
                                                                new JProperty("ImpactOnBusiness", p.ImpactOnBusiness),
                                                                new JProperty("Priority", p.Priority),
                                                                new JProperty("AssetIDs", p.AssetIDs),
                                                                new JProperty("NoOfPeopleAffected", p.NoOfPeopleAffected),
                                                                new JProperty("IsConfirmed", p.IsConfirmed),
                                                                new JProperty("RootCause", p.RootCause),
                                                                new JProperty("CorrectiveAction", p.CorrectiveAction),
                                                                new JProperty("PreventiveAction", p.PreventiveAction),
                                                                new JProperty("ActionCompletedByUserID", p.ActionCompletedByUserID),
                                                                new JProperty("ActionCompletedOn", p.ActionCompletedOn),
                                                                new JProperty("Remarks", p.Remarks),
                                                                new JProperty("IsActionConfirmed", p.IsActionConfirmed)
                                                                ))));

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Read Incident Years
        [HttpGet]
        [Route("ReadIncidentYears")]
        public HttpResponseMessage ReadIncidentYears()
        {
            try
            {
                DataTable dtDetails = new DataTable();
                List<string> IncidentYearsList = new List<string>();
                dtDetails = _BLLIncidentRegister.ReadIncidentYears();

                for (int i = 0; i < dtDetails.Rows.Count; i++)
                    IncidentYearsList.Add(dtDetails.Rows[i]["YearOfIncident"].ToString());

                JObject PEReport = new JObject(new JProperty("Success", 1),
                                                new JProperty("IncidentYearsList", new JArray(from p in IncidentYearsList
                                                                                              select new JObject(new JProperty("YearOfIncident", p.ToString())))));
                return Request.CreateResponse(HttpStatusCode.OK, PEReport);
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

        #region Update Incident
        [HttpPost]
        [Route("PostUpdateIncident")]
        public HttpResponseMessage PostUpdateIncident([FromBody] IncidentRegisterModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Register Incident"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objVar = new Result();

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Model Data");

                string strOutput = _BLLIncidentRegister.UpdateIncident(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objVar);
                }
                else
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objVar);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Delete Incident
        [HttpPost]
        [Route("PostDeleteIncident")]
        public HttpResponseMessage PostDeleteIncident([FromBody] IncidentRegisterModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Register Incident"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objVar = new Result();

                if (obj == null)
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Incident data is required");

                string strOutput = _BLLIncidentRegister.DeleteIncident(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objVar);
                }
                else
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objVar);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion 

        #region Update Action on Incident
        [HttpPost]
        [Route("PostUpdateActionOnIncident")]
        public HttpResponseMessage PostUpdateActionOnIncident([FromBody] IncidentRegisterModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Register Incident"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objVar = new Result();

                string strOutput = _BLLIncidentRegister.UpdateActionOnIncident(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objVar);
                }
                else
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objVar);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export the list of Incidents to Excel
        [HttpGet]
        [Route("ExportIncidentsListToExcel")]
        public HttpResponseMessage ExportIncidentsListToExcel(string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Report"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                string FileName = "Incidents List.xlsx";
                DataTable dtDetails = new DataTable();
                List<IncidentRegisterModel> incidentsList = new List<IncidentRegisterModel>();
                dtDetails = _BLLIncidentRegister.ReadAllIncidents();

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentRegisterModel objIncidentRegisterModel = new IncidentRegisterModel();
                        objIncidentRegisterModel.IncidentRegisterID = dtDetails.Rows[i]["IncidentRegisterID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentRegisterID"]) : 0;
                        objIncidentRegisterModel.IncidentNo = dtDetails.Rows[i]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[i]["IncidentNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["IncidentDate"] != DBNull.Value)
                            objIncidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[i]["IncidentDate"]);
                        objIncidentRegisterModel.IncidentTime = dtDetails.Rows[i]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[i]["IncidentTime"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentStatus = dtDetails.Rows[i]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[i]["IncidentStatus"].ToString().Trim() : "";
                        objIncidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentDescription = dtDetails.Rows[i]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[i]["IncidentDescription"].ToString().Trim() : "";
                        objIncidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[i]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[i]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                        objIncidentRegisterModel.ContactNo = dtDetails.Rows[i]["ContactNo"] != DBNull.Value ? dtDetails.Rows[i]["ContactNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.EmailID = dtDetails.Rows[i]["EmailID"] != DBNull.Value ? dtDetails.Rows[i]["EmailID"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentLocation = dtDetails.Rows[i]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[i]["IncidentLocation"].ToString().Trim() : "";
                        objIncidentRegisterModel.InformationAffected = dtDetails.Rows[i]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[i]["InformationAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.EquipmentAffected = dtDetails.Rows[i]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[i]["EquipmentAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[i]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[i]["ImpactOnBusiness"].ToString().Trim() : "";
                        objIncidentRegisterModel.Priority = dtDetails.Rows[i]["Priority"] != DBNull.Value ? dtDetails.Rows[i]["Priority"].ToString().Trim() : "";
                        objIncidentRegisterModel.AssetIDs = dtDetails.Rows[i]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[i]["AssetIDs"].ToString().Trim() : "";
                        objIncidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[i]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["NoOfPeopleAffected"]) : 0;
                        objIncidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsConfirmed"]);
                        objIncidentRegisterModel.RootCause = dtDetails.Rows[i]["RootCause"] != DBNull.Value ? dtDetails.Rows[i]["RootCause"].ToString().Trim() : "";
                        objIncidentRegisterModel.CorrectiveAction = dtDetails.Rows[i]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[i]["CorrectiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.PreventiveAction = dtDetails.Rows[i]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[i]["PreventiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[i]["ActionCompletedByUserName"] != DBNull.Value ? dtDetails.Rows[i]["ActionCompletedByUserName"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["ActionCompletedOn"] != DBNull.Value)
                            objIncidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[i]["ActionCompletedOn"]);
                        objIncidentRegisterModel.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        objIncidentRegisterModel.IsActionConfirmed = dtDetails.Rows[i]["IsActionConfirmed"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsActionConfirmed"]) : false;
                        incidentsList.Add(objIncidentRegisterModel);
                    }

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

                    #region Writing column headings and setting style
                    ws.Cells[0, 0].PutValue("S.No.");
                    ws.Cells[0, 1].PutValue("Incident ID");
                    ws.Cells[0, 2].PutValue("Incident No.");
                    ws.Cells[0, 3].PutValue("Incident Type");
                    ws.Cells[0, 4].PutValue("Incident Date");
                    ws.Cells[0, 5].PutValue("Incident Status");
                    ws.Cells[0, 6].PutValue("Department Resolving Incident");
                    ws.Cells[0, 7].PutValue("Incident Description");
                    ws.Cells[0, 8].PutValue("Name of Person Reporting Incident");
                    ws.Cells[0, 9].PutValue("Contact No.");
                    ws.Cells[0, 10].PutValue("Email ID");
                    ws.Cells[0, 11].PutValue("Incident Location");
                    ws.Cells[0, 12].PutValue("Information Affected");
                    ws.Cells[0, 13].PutValue("Equipment Affected");
                    ws.Cells[0, 14].PutValue("No. of People Affected");
                    ws.Cells[0, 15].PutValue("Impact On Business");
                    ws.Cells[0, 16].PutValue("Priority");
                    ws.Cells[0, 17].PutValue("Asset IDs");
                    ws.Cells[0, 18].PutValue("Is Incident Confirmed?");
                    ws.Cells[0, 19].PutValue("Root Cause");
                    ws.Cells[0, 20].PutValue("Corrective Action");
                    ws.Cells[0, 21].PutValue("Preventive Action");
                    ws.Cells[0, 22].PutValue("Action Completed By User Name");
                    ws.Cells[0, 23].PutValue("Action Completed on");
                    ws.Cells[0, 24].PutValue("Remarks");
                    ws.Cells[0, 25].PutValue("Is Action Confirmed?");

                    for (int c = 0; c <= 25; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (IncidentRegisterModel irmodel in incidentsList)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(row);
                        ws.Cells[row, 1].PutValue(irmodel.IncidentRegisterID);
                        ws.Cells[row, 2].PutValue(irmodel.IncidentNo);
                        ws.Cells[row, 3].PutValue(irmodel.IncidentType);
                        if (irmodel.IncidentDate != null)
                            ws.Cells[row, 4].PutValue(Convert.ToDateTime(irmodel.IncidentDate).ToString("dd-MMM-yyyy"));
                        ws.Cells[row, 5].PutValue(irmodel.IncidentStatus);
                        ws.Cells[row, 6].PutValue(irmodel.DepartmentResolvingIncident);
                        ws.Cells[row, 7].PutValue(irmodel.IncidentDescription);
                        ws.Cells[row, 8].PutValue(irmodel.NameOfPersonReportingIncident);
                        ws.Cells[row, 9].PutValue(irmodel.ContactNo);
                        ws.Cells[row, 10].PutValue(irmodel.EmailID);
                        ws.Cells[row, 11].PutValue(irmodel.IncidentLocation);
                        ws.Cells[row, 12].PutValue(irmodel.InformationAffected);
                        ws.Cells[row, 13].PutValue(irmodel.EquipmentAffected);
                        ws.Cells[row, 14].PutValue(irmodel.NoOfPeopleAffected);
                        ws.Cells[row, 15].PutValue(irmodel.ImpactOnBusiness);
                        ws.Cells[row, 16].PutValue(irmodel.Priority);
                        ws.Cells[row, 17].PutValue(irmodel.AssetIDs);
                        ws.Cells[row, 18].PutValue(irmodel.IsConfirmed == true?"Yes":"No");
                        if(irmodel.IsConfirmed)
                        {
                            ws.Cells[row, 19].PutValue(irmodel.RootCause);
                            ws.Cells[row, 20].PutValue(irmodel.CorrectiveAction);
                            ws.Cells[row, 21].PutValue(irmodel.PreventiveAction);
                            ws.Cells[row, 22].PutValue(irmodel.ActionCompletedByUserName);
                            if (irmodel.ActionCompletedOn != null)
                                ws.Cells[row, 23].PutValue(Convert.ToDateTime(irmodel.ActionCompletedOn).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 24].PutValue(irmodel.Remarks);
                            ws.Cells[row, 25].PutValue(irmodel.IsActionConfirmed==true?"Yes":"No");
                        }
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 7].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 12].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 16].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 19].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 20].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 21].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 24].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 25].SetStyle(styleCenterAlignData);
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
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Export the list of Incidents status count summary
        [HttpGet]
        [Route("ExportIncidentsStatusCountSummaryListToExcel")]
        public HttpResponseMessage ExportIncidentsStatusCountSummaryListToExcel(string YearOfIncident = null)
        {
            try
            {
                if (YearOfIncident != null && YearOfIncident.Trim().ToLower() == "all")
                    YearOfIncident = null;

                string FileName = "Incidents Status Count Summary.xlsx";
                DataTable dtDetails = new DataTable();
                List<IncidentCountSummary> lstIncidents = new List<IncidentCountSummary>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsCountSummaryByYear(YearOfIncident);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentCountSummary incidentCountSummary = new IncidentCountSummary();
                        incidentCountSummary.Department = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        incidentCountSummary.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        incidentCountSummary.PendingCount = dtDetails.Rows[i]["PendingCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["PendingCount"]) : 0;
                        incidentCountSummary.InProgressCount = dtDetails.Rows[i]["InProgressCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["InProgressCount"]) : 0;
                        incidentCountSummary.CompletedCount = dtDetails.Rows[i]["CompletedCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["CompletedCount"]) : 0;
                        incidentCountSummary.TotalCount = dtDetails.Rows[i]["TotalCount"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["TotalCount"]) : 0;
                        lstIncidents.Add(incidentCountSummary);
                    }

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

                    #region Writing column headings and setting style
                    ws.Cells[0, 0].PutValue("S.No.");
                    ws.Cells[0, 1].PutValue("Department");
                    ws.Cells[0, 2].PutValue("Incident Type");
                    ws.Cells[0, 3].PutValue("Pending Count");
                    ws.Cells[0, 4].PutValue("InProgress Count");
                    ws.Cells[0, 5].PutValue("Completed Count");
                    ws.Cells[0, 6].PutValue("Total Count");

                    for (int c = 0; c <= 6; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (IncidentCountSummary ics in lstIncidents)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(row);
                        ws.Cells[row, 1].PutValue(ics.Department);
                        ws.Cells[row, 2].PutValue(ics.IncidentType);
                        ws.Cells[row, 3].PutValue(ics.PendingCount);
                        ws.Cells[row, 4].PutValue(ics.InProgressCount);
                        ws.Cells[row, 5].PutValue(ics.CompletedCount);
                        ws.Cells[row, 6].PutValue(ics.TotalCount);
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 5].SetStyle(styleCenterAlignData);
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
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Export the list of Incidents By Year and Status to Excel
        [HttpGet]
        [Route("ExportIncidentsListByYearAndStatusToExcel")]
        public HttpResponseMessage ExportIncidentsListByYearAndStatusToExcel(string Department, string IncidentType, string Status, string YearOfIncident = null)
        {
            try
            {
                if (YearOfIncident != null && YearOfIncident.Trim().ToLower() == "all")
                    YearOfIncident = null;

                string FileName = "Incidents List By Year, Department, Incident Type and, Status.xlsx";
                DataTable dtDetails = new DataTable();
                List<IncidentRegisterModel> lstIncidents = new List<IncidentRegisterModel>();
                dtDetails = _BLLIncidentRegister.ReadIncidentsByIncidentYearAndStatus(Department, IncidentType, YearOfIncident, Status);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentRegisterModel objIncidentRegisterModel = new IncidentRegisterModel();
                        objIncidentRegisterModel.IncidentRegisterID = dtDetails.Rows[i]["IncidentRegisterID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentRegisterID"]) : 0;
                        objIncidentRegisterModel.IncidentNo = dtDetails.Rows[i]["IncidentNo"] != DBNull.Value ? dtDetails.Rows[i]["IncidentNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["IncidentDate"] != DBNull.Value)
                            objIncidentRegisterModel.IncidentDate = Convert.ToDateTime(dtDetails.Rows[i]["IncidentDate"]);
                        objIncidentRegisterModel.IncidentTime = dtDetails.Rows[i]["IncidentTime"] != DBNull.Value ? dtDetails.Rows[i]["IncidentTime"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentStatus = dtDetails.Rows[i]["IncidentStatus"] != DBNull.Value ? dtDetails.Rows[i]["IncidentStatus"].ToString().Trim() : "";
                        objIncidentRegisterModel.DepartmentResolvingIncident = dtDetails.Rows[i]["Department"] != DBNull.Value ? dtDetails.Rows[i]["Department"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentDescription = dtDetails.Rows[i]["IncidentDescription"] != DBNull.Value ? dtDetails.Rows[i]["IncidentDescription"].ToString().Trim() : "";
                        objIncidentRegisterModel.NameOfPersonReportingIncident = dtDetails.Rows[i]["NameOfPersonReportingIncident"] != DBNull.Value ? dtDetails.Rows[i]["NameOfPersonReportingIncident"].ToString().Trim() : "";
                        objIncidentRegisterModel.ContactNo = dtDetails.Rows[i]["ContactNo"] != DBNull.Value ? dtDetails.Rows[i]["ContactNo"].ToString().Trim() : "";
                        objIncidentRegisterModel.EmailID = dtDetails.Rows[i]["EmailID"] != DBNull.Value ? dtDetails.Rows[i]["EmailID"].ToString().Trim() : "";
                        objIncidentRegisterModel.IncidentLocation = dtDetails.Rows[i]["IncidentLocation"] != DBNull.Value ? dtDetails.Rows[i]["IncidentLocation"].ToString().Trim() : "";
                        objIncidentRegisterModel.InformationAffected = dtDetails.Rows[i]["InformationAffected"] != DBNull.Value ? dtDetails.Rows[i]["InformationAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.EquipmentAffected = dtDetails.Rows[i]["EquipmentAffected"] != DBNull.Value ? dtDetails.Rows[i]["EquipmentAffected"].ToString().Trim() : "";
                        objIncidentRegisterModel.ImpactOnBusiness = dtDetails.Rows[i]["ImpactOnBusiness"] != DBNull.Value ? dtDetails.Rows[i]["ImpactOnBusiness"].ToString().Trim() : "";
                        objIncidentRegisterModel.Priority = dtDetails.Rows[i]["Priority"] != DBNull.Value ? dtDetails.Rows[i]["Priority"].ToString().Trim() : "";
                        objIncidentRegisterModel.AssetIDs = dtDetails.Rows[i]["AssetIDs"] != DBNull.Value ? dtDetails.Rows[i]["AssetIDs"].ToString().Trim() : "";
                        objIncidentRegisterModel.NoOfPeopleAffected = dtDetails.Rows[i]["NoOfPeopleAffected"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["NoOfPeopleAffected"]) : 0;
                        if (dtDetails.Rows[i]["IsConfirmed"] != DBNull.Value)
                            objIncidentRegisterModel.IsConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsConfirmed"]);
                        objIncidentRegisterModel.RootCause = dtDetails.Rows[i]["RootCause"] != DBNull.Value ? dtDetails.Rows[i]["RootCause"].ToString().Trim() : "";
                        objIncidentRegisterModel.CorrectiveAction = dtDetails.Rows[i]["CorrectiveAction"] != DBNull.Value ? dtDetails.Rows[i]["CorrectiveAction"].ToString().Trim() : "";
                        objIncidentRegisterModel.PreventiveAction = dtDetails.Rows[i]["PreventiveAction"] != DBNull.Value ? dtDetails.Rows[i]["PreventiveAction"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["ActionCompletedByUserID"] != DBNull.Value)
                        {
                            objIncidentRegisterModel.ActionCompletedByUserID = Convert.ToInt32(dtDetails.Rows[i]["ActionCompletedByUserID"]);
                            objIncidentRegisterModel.ActionCompletedByUserName = dtDetails.Rows[i]["ActionCompletedByUserName"].ToString().Trim();
                        }
                        if (dtDetails.Rows[i]["ActionCompletedOn"] != DBNull.Value)
                            objIncidentRegisterModel.ActionCompletedOn = Convert.ToDateTime(dtDetails.Rows[i]["ActionCompletedOn"]);
                        objIncidentRegisterModel.Remarks = dtDetails.Rows[i]["Remarks"] != DBNull.Value ? dtDetails.Rows[i]["Remarks"].ToString().Trim() : "";
                        if (dtDetails.Rows[i]["IsActionConfirmed"] != DBNull.Value)
                            objIncidentRegisterModel.IsActionConfirmed = Convert.ToBoolean(dtDetails.Rows[i]["IsActionConfirmed"]);
                        lstIncidents.Add(objIncidentRegisterModel);
                    }

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

                    #region Writing column headings and setting style
                    ws.Cells[0, 0].PutValue("S.No.");
                    ws.Cells[0, 1].PutValue("Incident ID");
                    ws.Cells[0, 2].PutValue("Incident No.");
                    ws.Cells[0, 3].PutValue("Incident Type");
                    ws.Cells[0, 4].PutValue("Incident Date");
                    ws.Cells[0, 5].PutValue("Incident Status");
                    ws.Cells[0, 6].PutValue("Department Resolving Incident");
                    ws.Cells[0, 7].PutValue("Incident Description");
                    ws.Cells[0, 8].PutValue("Name of Person Reporting Incident");
                    ws.Cells[0, 9].PutValue("Contact No.");
                    ws.Cells[0, 10].PutValue("Email ID");
                    ws.Cells[0, 11].PutValue("Incident Location");
                    ws.Cells[0, 12].PutValue("Information Affected");
                    ws.Cells[0, 13].PutValue("Equipment Affected");
                    ws.Cells[0, 14].PutValue("No. of People Affected");
                    ws.Cells[0, 15].PutValue("Impact On Business");
                    ws.Cells[0, 16].PutValue("Priority");
                    ws.Cells[0, 17].PutValue("Asset IDs");
                    ws.Cells[0, 18].PutValue("Is Incident Confirmed?");
                    ws.Cells[0, 19].PutValue("Root Cause");
                    ws.Cells[0, 20].PutValue("Corrective Action");
                    ws.Cells[0, 21].PutValue("Preventive Action");
                    ws.Cells[0, 22].PutValue("Action Completed By User Name");
                    ws.Cells[0, 23].PutValue("Action Completed on");
                    ws.Cells[0, 24].PutValue("Remarks");
                    ws.Cells[0, 25].PutValue("Is Action Confirmed?");

                    for (int c = 0; c <= 25; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (IncidentRegisterModel irmodel in lstIncidents)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(row);
                        ws.Cells[row, 1].PutValue(irmodel.IncidentRegisterID);
                        ws.Cells[row, 2].PutValue(irmodel.IncidentNo);
                        ws.Cells[row, 3].PutValue(irmodel.IncidentType);
                        if (irmodel.IncidentDate != null)
                            ws.Cells[row, 4].PutValue(Convert.ToDateTime(irmodel.IncidentDate).ToString("dd-MMM-yyyy"));
                        ws.Cells[row, 5].PutValue(irmodel.IncidentStatus);
                        ws.Cells[row, 6].PutValue(irmodel.DepartmentResolvingIncident);
                        ws.Cells[row, 7].PutValue(irmodel.IncidentDescription);
                        ws.Cells[row, 8].PutValue(irmodel.NameOfPersonReportingIncident);
                        ws.Cells[row, 9].PutValue(irmodel.ContactNo);
                        ws.Cells[row, 10].PutValue(irmodel.EmailID);
                        ws.Cells[row, 11].PutValue(irmodel.IncidentLocation);
                        ws.Cells[row, 12].PutValue(irmodel.InformationAffected);
                        ws.Cells[row, 13].PutValue(irmodel.EquipmentAffected);
                        ws.Cells[row, 14].PutValue(irmodel.NoOfPeopleAffected);
                        ws.Cells[row, 15].PutValue(irmodel.ImpactOnBusiness);
                        ws.Cells[row, 16].PutValue(irmodel.Priority);
                        ws.Cells[row, 17].PutValue(irmodel.AssetIDs);
                        ws.Cells[row, 18].PutValue(irmodel.IsConfirmed == true ? "Yes" : "No");
                        if (irmodel.IsConfirmed)
                        {
                            ws.Cells[row, 19].PutValue(irmodel.RootCause);
                            ws.Cells[row, 20].PutValue(irmodel.CorrectiveAction);
                            ws.Cells[row, 21].PutValue(irmodel.PreventiveAction);
                            ws.Cells[row, 22].PutValue(irmodel.ActionCompletedByUserName);
                            if (irmodel.ActionCompletedOn != null)
                                ws.Cells[row, 23].PutValue(Convert.ToDateTime(irmodel.ActionCompletedOn).ToString("dd-MMM-yyyy"));
                            ws.Cells[row, 24].PutValue(irmodel.Remarks);
                            ws.Cells[row, 25].PutValue(irmodel.IsActionConfirmed == true ? "Yes" : "No");
                        }
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 7].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 8].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 9].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 10].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 11].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 12].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 13].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 14].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 15].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 16].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 17].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 18].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 19].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 20].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 21].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 22].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 23].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 24].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 25].SetStyle(styleCenterAlignData);
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
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Download Help Document file
        [HttpGet]
        [Route("DownloadHelpDocument")]
        public HttpResponseMessage DownloadHelpDocument()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileName = "Incident Report - User Guide.pptx";

                //Set the Help File Path
                string filePath = HttpContext.Current.Server.MapPath("~/HelpDocs/") + FileName;

                //Check whether File exists.
                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.NotFound, "File not found");

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
    }
}
