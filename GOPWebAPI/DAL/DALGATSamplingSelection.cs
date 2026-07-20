using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALGATSamplingSelection
    {
        private readonly string _connectionString;

        public DALGATSamplingSelection(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Check Duplicate Material No. exists
        public string CheckDuplicateMaterialNoExists(string sqlTableName)
        {
            string strResult = string.Empty;

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand("spGATSSCheckDuplicateMaterialNo", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        #region Add input Parameters with values
                        command.CommandTimeout = 500;
                        command.Parameters.AddWithValue("@TableName", sqlTableName);
                        // Output parameter to capture the message
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);
                        #endregion

                        connection.Open();
                        command.ExecuteNonQuery();
                        strResult = Convert.ToString(command.Parameters["@Result"].Value);
                        connection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strResult = "Error: " + ex.Message;
            }

            return strResult;
        }
        #endregion

        #region Select And Fetch Random Rows
        public DataTable SelectAndFetchRandomRows(string sqlTableName, int SamplingSelectionPercentage)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATSamplingSelectionGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        cmd.Parameters.AddWithValue("@TableName", sqlTableName);
                        cmd.Parameters.AddWithValue("@SamplingPercentage", SamplingSelectionPercentage);
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

        #region Fetch Items Count for each Noun
        public DataTable FetchItemsCountForEachNoun(string sqlTableName)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATSamplingSelectionGetNounCount", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        cmd.Parameters.AddWithValue("@TableName", sqlTableName);
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

        #region Items Count for each Noun from randomly selected Items
        public DataTable FetchItemsCountForEachNounFromRandomlySelectedItems(string sqlTableName, int SamplingSelectionPercentage)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATSSGetSelectedNounCount", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        cmd.Parameters.AddWithValue("@TableName", sqlTableName);
                        cmd.Parameters.AddWithValue("@SamplingPercentage", SamplingSelectionPercentage);
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