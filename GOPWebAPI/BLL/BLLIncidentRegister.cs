using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;
using GOPWebAPI.Models.Incident_Report_Models;

namespace GOPWebAPI.BLL
{
    public class BLLIncidentRegister
    {
        private readonly string _connectionString;
        private readonly DALIncidentRegister _DALIncidentRegister;
        public BLLIncidentRegister(string connectionString)
        {
            _connectionString = connectionString;
            _DALIncidentRegister = new DALIncidentRegister(_connectionString);
        }

        #region Register an Incident
        public string RegisterIncident(IncidentRegisterModel objVar)
        {
            return _DALIncidentRegister.RegisterIncident(objVar);
        }
        #endregion

        #region Read Incidents Unique Search Values by Search Field
        public DataTable ReadIncidentsUniqueSearchValuesBySearchField(string SearchField)
        {
            return _DALIncidentRegister.ReadIncidentsUniqueSearchValuesBySearchField(SearchField);
        }
        #endregion

        #region Read All Incidents
        public DataTable ReadAllIncidents()
        {
            return _DALIncidentRegister.ReadAllIncidents();
        }
        #endregion

        #region Read Incident By Id
        public DataTable ReadIncidentById(long incidentRegisterID)
        {
            return _DALIncidentRegister.ReadIncidentById(incidentRegisterID);
        }
        #endregion

        #region Read Incident Departments Affected By Id
        public DataTable ReadDepartmentsAffectedById(long incidentRegisterID)
        {
            return _DALIncidentRegister.ReadDepartmentsAffectedById(incidentRegisterID);
        }
        #endregion

        #region Read Incidents By Search Field and Search Value
        public DataTable ReadIncidentsBySearchFieldAndSearchValue(string SearchField, string SearchValue)
        {
            return _DALIncidentRegister.ReadIncidentsBySearchFieldAndSearchValue(SearchField, SearchValue);
        }
        #endregion

        #region Read Incidents Count Summary with status and each incident type for selected year and / or all
        public DataTable ReadIncidentsCountSummaryByYear(string YearOfIncident = null)
        {
            return _DALIncidentRegister.ReadIncidentsCountSummaryByYear(YearOfIncident);
        }
        #endregion

        #region Read Incidents By Year and Status
        public DataTable ReadIncidentsByIncidentYearAndStatus(string Department, string IncidentType, string Status, string YearOfIncident = null)
        {
            return _DALIncidentRegister.ReadIncidentsByIncidentYearAndStatus(Department, IncidentType, YearOfIncident, Status);
        }
        #endregion

        #region Update Incident
        public string UpdateIncident(IncidentRegisterModel objVar)
        {
            return _DALIncidentRegister.UpdateIncident(objVar);
        }
        #endregion

        #region Delete Incident
        public string DeleteIncident(IncidentRegisterModel objVar)
        {
            return _DALIncidentRegister.DeleteIncident(objVar);
        }
        #endregion

        #region Update Action on Incident
        public string UpdateActionOnIncident(IncidentRegisterModel objVar)
        {
            return _DALIncidentRegister.UpdateActionOnIncident(objVar);
        }
        #endregion

        #region Read Incident Years
        public DataTable ReadIncidentYears()
        {
            return _DALIncidentRegister.ReadIncidentYears();
        }
        #endregion
    }
}