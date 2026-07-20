using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;
using GOPWebAPI.Helpers;

namespace GOPWebAPI.BLL
{
    public class BLLProduction
    {
        private readonly string _connectionString;
        private readonly DALProduction _DALProduction;

        public BLLProduction(string connectionString)
        {
            _connectionString = connectionString;
            _DALProduction = new DALProduction(_connectionString);
        }

        #region Read Project Overall Status all results
        public DataSet ReadProjectOverallStatusAllResults(string CustomerCode, string ProjectCode, string BatchNo, string Status)
        {
            return _DALProduction.ReadProjectOverallStatusAllResults(CustomerCode, ProjectCode, BatchNo, Status);
        }
        #endregion

        #region Read Date Range Status
        public DataTable ReadDateRangeStatus(DateTime FromDate, DateTime ToDate, string Status)
        {
            return _DALProduction.ReadDateRangeStatus(FromDate, ToDate, Status);
        }
        #endregion

        #region User Based Status Report
        public DataTable ReadProductionAndQCUserNames()
        {
            return _DALProduction.ReadProductionAndQCUserNames();
        }

        public DataTable ReadUserBasedStatus(string UserName, string Status)
        {
            return _DALProduction.ReadUserBasedStatus(UserName, Status);
        }
        #endregion

        #region Project Level Quality Report
        #region Project Level Quality Report Count Stats
        public DataTable ReadProjectLevelQualityCountStats(string CustomerCode, string ProjectCode, string BatchNo)
        {
            return _DALProduction.ReadProjectLevelQualityCountStats(CustomerCode, ProjectCode, BatchNo);
        }
        #endregion

        #region Project Level Quality Report SKU Details
        public DataTable ReadProjectLevelQualityReportSKUDetails(string CustomerCode, string ProjectCode, string BatchNo, string Status)
        {
            return _DALProduction.ReadProjectLevelQualityReportSKUDetails(CustomerCode, ProjectCode, BatchNo, Status);
        }
        #endregion

        #endregion

        #region Resource Level Quality Report
        #region Resource Level Quality Report Count Stats
        public DataTable ReadResourceLevelQualityCountStats(string UserName, string CustomerCode, string ProjectCode, string BatchNo, DateTime? FromDate, DateTime? ToDate)
        {
            return _DALProduction.ReadResourceLevelQualityCountStats(UserName, CustomerCode, ProjectCode, BatchNo, FromDate, ToDate);
        }
        #endregion

        #region Resource Level Quality Report SKU Details
        public DataTable ReadResourceLevelQualityReportSKUDetails(string UserName, string CustomerCode, string ProjectCode, string BatchNo, DateTime? FromDate, DateTime? ToDate, string Status)
        {
            return _DALProduction.ReadResourceLevelQualityReportSKUDetails(UserName, CustomerCode, ProjectCode, BatchNo, FromDate, ToDate, Status);
        }
        #endregion
        #endregion

        #region Read MRO Ref. DB Part Details by Part No.
        public DataTable ReadMRORefDBPartDetailsByPartNo(string PartNo)
        {
            return _DALProduction.ReadMRORefDBPartDetailsByPartNo(PartNo);
        }
        #endregion
    }
}