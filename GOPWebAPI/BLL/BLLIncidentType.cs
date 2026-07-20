using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;
using GOPWebAPI.Models.Incident_Report_Models;

namespace GOPWebAPI.BLL
{
    public class BLLIncidentType
    {
        private readonly string _connectionString;
        private readonly DALIncidentType _DALIncidentType;
        public BLLIncidentType(string connectionString)
        {
            _connectionString = connectionString;
            _DALIncidentType = new DALIncidentType(_connectionString);
        }

        #region Add Incident Type
        public string AddIncidentType(IncidentTypeModel objVar)
        {
            return _DALIncidentType.AddIncidentType(objVar);
        }
        #endregion

        #region Read All Incident Types
        public DataTable ReadAllIncidentTypes()
        {
            return _DALIncidentType.ReadAllIncidentTypes();
        }
        #endregion

        #region Read Incident Type By Id
        public DataTable ReadIncidentTypeById(int IncidentTypeID)
        {
            return _DALIncidentType.ReadIncidentTypeById(IncidentTypeID);
        }
        #endregion

        #region Update Incident Type
        public string UpdateIncidentType(IncidentTypeModel objVar)
        {
            return _DALIncidentType.UpdateIncidentType(objVar);
        }
        #endregion

        #region Delete Incident Type
        public string DeleteIncidentType(IncidentTypeModel objVar)
        {
            return _DALIncidentType.DeleteIncidentType(objVar);
        }
        #endregion
    }
}