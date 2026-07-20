using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.Incident_Report_Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALMarketingDocs
    {
        private readonly string _connectionString;

        public DALMarketingDocs(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Upload Marketing Document
        public string UploadMarketingDocument(MarketingDocModel model)
        {
            string strResult = string.Empty;

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand("spMarketingDocs", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        #region Add Parameters with values
                        command.CommandTimeout = 500;
                        command.Parameters.AddWithValue("@FileName", model.UserUploadedFileName);
                        command.Parameters.AddWithValue("@Domain", model.Domain);
                        command.Parameters.AddWithValue("@DocType", model.DocType);
                        command.Parameters.AddWithValue("@UserID", model.UserID);
                        command.Parameters.AddWithValue("@Mode", 1);    //1--Upload, 2-Update, 3-Delete, 4-Read All, 5-Read By Id
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

                return strResult;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strResult = "Error: " + ex.Message;
                return strResult;
            }
        }
        #endregion

        #region Read All Marketing Documents from input domain and doc type
        public DataTable ReadAllMarketingDocuments(string Domain, string DocType)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMarketingDocs", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Domain", Domain);
                        cmd.Parameters.AddWithValue("@DocType", DocType);
                        cmd.Parameters.AddWithValue("@Mode", 4);    //flag to read all
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

        #region Read Marketing Document Details by id
        public DataTable ReadMarketingDocumentById(int id)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMarketingDocs", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Id", id);
                        cmd.Parameters.AddWithValue("@Mode", 5);
                        cmd.CommandTimeout = 500;
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

        #region Delete Marketing Document
        public string DeleteMarketingDocument(int id, string UserID)
        {
            string strOutput = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand("spMarketingDocs", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        #region Add Parameters with values
                        command.CommandTimeout = 500;
                        command.Parameters.AddWithValue("@id", id);
                        command.Parameters.AddWithValue("@UserID", UserID);
                        command.Parameters.AddWithValue("@Mode", 3);
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);
                        #endregion

                        connection.Open();
                        command.ExecuteNonQuery();
                        strOutput = Convert.ToString(command.Parameters["@Result"].Value);
                        connection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strOutput = "Error: " + ex.Message;
                return strOutput;
            }
            return strOutput;
        }
        #endregion

    }
}