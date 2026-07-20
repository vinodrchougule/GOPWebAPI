using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALGATSTPInputFormatConversion
    {
        private readonly string _connectionString;

        public DALGATSTPInputFormatConversion(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Fetch Input Format Conversion Details
        public DataTable FetchInputFormatConversionDetails(string InputFileTableName, string TaxonomyFileTableName)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATSTPIFCUpdateInputFileAndGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.AddWithValue("@InputFileTableName", InputFileTableName);
                        cmd.Parameters.AddWithValue("@TaxonomyFileTableName", TaxonomyFileTableName);
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dtDetails;
        }
        #endregion
    }
}