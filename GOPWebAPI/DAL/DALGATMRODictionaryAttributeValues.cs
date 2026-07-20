using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALGATMRODictionaryAttributeValues
    {
        private readonly string _connectionString;

        public DALGATMRODictionaryAttributeValues(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Fetch MRO Dictionary Attribute Values from the version
        public DataTable FetchMRODictionaryAttributeValues(string VersionNameOrNo)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATNounModifierAttributeValues", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
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