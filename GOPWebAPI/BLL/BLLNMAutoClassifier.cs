using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;

namespace GOPWebAPI.BLL
{
    public class BLLNMAutoClassifier
    {
        private readonly string _connectionString;
        private readonly DALNMAutoClassifier _DALNMAutoClassifier;

        public BLLNMAutoClassifier(string connectionString)
        {
            _connectionString = connectionString;
            _DALNMAutoClassifier = new DALNMAutoClassifier(_connectionString);
        }

        #region Process NM Auto Classifier
        public DataTable ProcessNMAutoClassifier(string InputFileTableName, string AbbreviationFileTableName, string StdNounModifierFileTableName)
        {
            return _DALNMAutoClassifier.ProcessNMAutoClassifier(InputFileTableName, AbbreviationFileTableName, StdNounModifierFileTableName);
        }
        #endregion
    }
}