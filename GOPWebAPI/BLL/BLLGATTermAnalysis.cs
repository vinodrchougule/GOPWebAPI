using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGATTermAnalysis
    {
        private readonly string _connectionString;
        private readonly DALGATTermAnalysis _DALGATTermAnalysis;

        public BLLGATTermAnalysis(string connectionString)
        {
            _connectionString = connectionString;
            _DALGATTermAnalysis = new DALGATTermAnalysis(_connectionString);
        }

        #region Get the Term Frequency data
        public DataTable GetTermFrequencyData(string descSqlTableName, char Delimiter)
        {
            return _DALGATTermAnalysis.GetTermFrequencyData(descSqlTableName, Delimiter);
        }
        #endregion

        #region Get the data for Term Attributes
        public DataTable GetTermAttributesData(string descSqlTableName)
        {
            return _DALGATTermAnalysis.GetTermAttributesData(descSqlTableName);
        }
        #endregion

        #region Get the data for Repeated Words
        public DataTable GetRepeatedWordsData(string descSqlTableName)
        {
            return _DALGATTermAnalysis.GetRepeatedWordsData(descSqlTableName);
        }
        #endregion

    }
}