using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLMapUNSPSCLevels
    {
        private readonly string _connectionString;
        private readonly DALMapUNSPSCLevels _DALMapUNSPSCLevels;
        public BLLMapUNSPSCLevels(string connectionString)
        {
            _connectionString = connectionString;
            _DALMapUNSPSCLevels = new DALMapUNSPSCLevels();
        }

        #region  Map UNSPSC Levels and return the output
        public DataTable MapUNSPSCLevels(string sqlTableName, string UNSPSCVersion)
        {
            return _DALMapUNSPSCLevels.MapUNSPSCLevels(sqlTableName, UNSPSCVersion);
        }
        #endregion
    }
}