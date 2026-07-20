using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGATMRODictionaryAttributeValues
    {
        private readonly string _connectionString;
        private readonly DALGATMRODictionaryAttributeValues _DALGATMRODictionaryAttributeValues;

        public BLLGATMRODictionaryAttributeValues(string connectionString)
        {
            _connectionString = connectionString;
            _DALGATMRODictionaryAttributeValues = new DALGATMRODictionaryAttributeValues(_connectionString);
        }

        #region Fetch MRO Dictionary Attribute Values from the version
        public DataTable FetchMRODictionaryAttributeValues(string VersionNameOrNo)
        {
            return _DALGATMRODictionaryAttributeValues.FetchMRODictionaryAttributeValues(VersionNameOrNo);
        }
        #endregion
    }
}