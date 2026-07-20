using GOPWebAPI.DAL;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLConfirmDuplicateSKUs
    {
        private readonly string _connectionString;
        private readonly DALConfirmDuplicateSKUs _DALConfirmDuplicateSKUs;

        public BLLConfirmDuplicateSKUs(string connectionString)
        {
            _connectionString = connectionString;
            _DALConfirmDuplicateSKUs = new DALConfirmDuplicateSKUs(_connectionString);
        }

        #region Check Is CIF SQL table exists
        public string IsCIFSQLTableExists(string CustomerCode, string ProjectCode, string BatchNo)
        {
            return _DALConfirmDuplicateSKUs.IsCIFSQLTableExists(CustomerCode, ProjectCode, BatchNo);
        }
        #endregion

        #region Create CIF SQL Table in database
        public string CreateCIFSQLTable(string CustomerCode, string ProjectCode, string BatchNo, string CreateTableSQLString)
        {
            return _DALConfirmDuplicateSKUs.CreateCIFSQLTable(CustomerCode, ProjectCode, BatchNo, CreateTableSQLString);
        }
        #endregion

        #region Update selected SKUs as Duplicates in CIF table on selected columns
        public string UpdateSelectedSKUsAsDuplicates(ConfirmDuplicateSKUsModel model)
        {
            return _DALConfirmDuplicateSKUs.UpdateSelectedSKUsAsDuplicates(model);
        }
        #endregion

        #region Fetch Customer Input File SKUs updated in SQL table
        public DataTable FetchCIFSKUsFromTable(ConfirmDuplicateSKUsModel model)
        {
            return _DALConfirmDuplicateSKUs.FetchCIFSKUsFromTable(model);
        }
        #endregion
    }
}