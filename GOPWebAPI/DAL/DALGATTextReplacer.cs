using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALGATTextReplacer
    {
        private readonly string _connectionString;

        public DALGATTextReplacer(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Replace Description with Terms based on Delimiter and Get the data
        public DataTable ReplaceDescriptionAndGetData(string descSqlTableName, string termsSqlTableName, string Delimiter)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    if (Delimiter.Trim().ToUpper() == "COMMA & SPACE")
                    {
                        using (SqlCommand cmd = new SqlCommand("spGATTRCSDescriptionGetData", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 0;
                            cmd.Parameters.AddWithValue("@DescTableName", descSqlTableName);
                            cmd.Parameters.AddWithValue("@TermsTableName", termsSqlTableName);
                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(dtDetails);
                            con.Close();
                        }
                    }
                    else
                    {
                        using (SqlCommand cmd = new SqlCommand("spGATTRDescriptionGetData", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 0;
                            cmd.Parameters.AddWithValue("@DescTableName", descSqlTableName);
                            cmd.Parameters.AddWithValue("@TermsTableName", termsSqlTableName);
                            cmd.Parameters.AddWithValue("@Delimiter", Delimiter);
                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(dtDetails);
                            con.Close();
                        }
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

        #region Get the terms with count
        public DataTable GetTermsWithCount(string descSqlTableName, string termsSqlTableName, string Delimiter)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    if (Delimiter.Trim().ToUpper() == "COMMA & SPACE")
                    {
                        using (SqlCommand cmd = new SqlCommand("spGATTRCSTermsCountGetData", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 0;
                            cmd.Parameters.AddWithValue("@DescTableName", descSqlTableName);
                            cmd.Parameters.AddWithValue("@TermsTableName", termsSqlTableName);
                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(dtDetails);
                            con.Close();
                        }
                    }
                    else
                    {
                        using (SqlCommand cmd = new SqlCommand("spGATTRTermsCountGetData", con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.CommandTimeout = 0;
                            cmd.Parameters.AddWithValue("@DescTableName", descSqlTableName);
                            cmd.Parameters.AddWithValue("@TermsTableName", termsSqlTableName);
                            cmd.Parameters.AddWithValue("@Delimiter", Delimiter);
                            con.Open();
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(dtDetails);
                            con.Close();
                        }
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