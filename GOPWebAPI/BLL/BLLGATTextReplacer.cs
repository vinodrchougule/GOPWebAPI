using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGATTextReplacer
    {
        private readonly string _connectionString;
        private readonly DALGATTextReplacer _DALGATTextReplacer;

        public BLLGATTextReplacer(string connectionString)
        {
            _connectionString = connectionString;
            _DALGATTextReplacer = new DALGATTextReplacer(_connectionString);
        }

        #region Replace Description with Terms based on Delimiter and Get the data
        public DataTable ReplaceDescriptionAndGetData(string descSqlTableName, string termsSqlTableName, string Delimiter)
        {
            return _DALGATTextReplacer.ReplaceDescriptionAndGetData(descSqlTableName, termsSqlTableName, Delimiter);
        }
        #endregion

        #region Get the terms with count
        public DataTable GetTermsWithCount(string descSqlTableName, string termsSqlTableName, string Delimiter)
        {
            return _DALGATTextReplacer.GetTermsWithCount(descSqlTableName,termsSqlTableName, Delimiter);
        }
        #endregion
    }
}