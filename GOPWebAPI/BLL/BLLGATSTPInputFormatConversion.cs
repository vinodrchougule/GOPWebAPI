using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGATSTPInputFormatConversion
    {
        private readonly string _connectionString;
        private readonly DALGATSTPInputFormatConversion _DALSTPInputFormatConversion;

        public BLLGATSTPInputFormatConversion(string connectionString)
        {
            _connectionString = connectionString;
            _DALSTPInputFormatConversion = new DALGATSTPInputFormatConversion(_connectionString);
        }

        #region Fetch Input Format Conversion Details
        public DataTable FetchInputFormatConversionDetails(string InputFileTableName, string TaxonomyFileTableName)
        {
            return _DALSTPInputFormatConversion.FetchInputFormatConversionDetails(InputFileTableName, TaxonomyFileTableName);
        }
        #endregion
    }
}