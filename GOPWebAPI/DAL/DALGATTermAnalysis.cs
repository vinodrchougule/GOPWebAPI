using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALGATTermAnalysis
    {
        private readonly string _connectionString;

        public DALGATTermAnalysis(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Get the Term Frequency data
        public DataTable GetTermFrequencyData(string descSqlTableName, char Delimiter)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATTATFGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.AddWithValue("@TableName", descSqlTableName);
                        cmd.Parameters.AddWithValue("@Delimiter", Delimiter);
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

        #region Get the data for Term Attributes
        public DataTable GetTermAttributesData(string descSqlTableName)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATTATAGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.AddWithValue("@TableName", descSqlTableName);
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

        #region Get the data for Repeated Words
        public DataTable GetRepeatedWordsData(string descSqlTableName)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATTARWGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 0;
                        cmd.Parameters.AddWithValue("@TableName", descSqlTableName);
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