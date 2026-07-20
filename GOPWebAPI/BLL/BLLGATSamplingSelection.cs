using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGATSamplingSelection
    {
        private readonly string _connectionString;
        private readonly DALGATSamplingSelection _DALGATSamplingSelection;

        public BLLGATSamplingSelection(string connectionString)
        {
            _connectionString = connectionString;
            _DALGATSamplingSelection = new DALGATSamplingSelection(_connectionString);
        }

        #region Check Duplicate Material No. Exists
        public string CheckDuplicateMaterialNoExists(string sqlTableName)
        {
            return _DALGATSamplingSelection.CheckDuplicateMaterialNoExists(sqlTableName);
        }
        #endregion

        #region Select And Fetch Random Rows
        public DataTable SelectAndFetchRandomRows(string sqlTableName, int SamplingSelectionPercentage)
        {
            return _DALGATSamplingSelection.SelectAndFetchRandomRows(sqlTableName, SamplingSelectionPercentage);
        }
        #endregion


        #region Fetch Items Count for each Noun
        public DataTable FetchItemsCountForEachNoun(string sqlTableName)
        {
            return _DALGATSamplingSelection.FetchItemsCountForEachNoun(sqlTableName);
        }
        #endregion

        #region Items Count for each Noun from randomly selected Items
        public DataTable FetchItemsCountForEachNounFromRandomlySelectedItems(string sqlTableName, int SamplingSelectionPercentage)
        {
            return _DALGATSamplingSelection.FetchItemsCountForEachNounFromRandomlySelectedItems(sqlTableName, SamplingSelectionPercentage);
        }
        #endregion
    }
}