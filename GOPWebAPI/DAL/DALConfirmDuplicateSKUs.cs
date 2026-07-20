using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml.Serialization;
using System.Xml;

namespace GOPWebAPI.DAL
{
    public class DALConfirmDuplicateSKUs
    {
        private readonly string _connectionString;

        public DALConfirmDuplicateSKUs(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Check Is CIF SQL table exists
        public string IsCIFSQLTableExists(string CustomerCode, string ProjectCode, string BatchNo)
        {
            string strOutput = string.Empty;

            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmdCreateSQLTable = new SqlCommand("spCIFSQLTableExists", conn))
                    {
                        cmdCreateSQLTable.CommandType = CommandType.StoredProcedure;
                        cmdCreateSQLTable.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmdCreateSQLTable.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmdCreateSQLTable.Parameters.AddWithValue("@BatchNo", BatchNo);
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        cmdCreateSQLTable.Parameters.Add(outMsgParam);

                        conn.Open();
                        cmdCreateSQLTable.ExecuteNonQuery();
                        strOutput = Convert.ToString(cmdCreateSQLTable.Parameters["@Result"].Value);
                        conn.Close();
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

        #region Create CIF SQL Table in database
        public string CreateCIFSQLTable(string CustomerCode, string ProjectCode, string BatchNo, string CreateTableSQLString)
        {
            string strOutput = string.Empty;

            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmdCreateSQLTable = new SqlCommand("spCreateCIFSQLTable", conn))
                    {
                        cmdCreateSQLTable.CommandType = CommandType.StoredProcedure;
                        cmdCreateSQLTable.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmdCreateSQLTable.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmdCreateSQLTable.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmdCreateSQLTable.Parameters.AddWithValue("@CreateTableSQLString", CreateTableSQLString);
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        cmdCreateSQLTable.Parameters.Add(outMsgParam);

                        conn.Open();
                        cmdCreateSQLTable.ExecuteNonQuery();
                        strOutput = Convert.ToString(cmdCreateSQLTable.Parameters["@Result"].Value);
                        conn.Close();
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

        #region Update selected SKUs as Duplicates in CIF table on selected columns
        public string UpdateSelectedSKUsAsDuplicates(ConfirmDuplicateSKUsModel model)
        {
            string strOutput = string.Empty;

            try
            {
                #region CIF IDs XML Serialization
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<CIFID>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, model.CIFIDs);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmdUpdateSelectedSKUsAsDuplicates = new SqlCommand("spProjectDuplicateSKUsFromCIF", conn))
                    {
                        cmdUpdateSelectedSKUsAsDuplicates.CommandType = CommandType.StoredProcedure;
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@SelectedColumns", model.SelectedColumns);
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@CIFIDs", sqlXml);
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.AddWithValue("@Mode", 1);
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        cmdUpdateSelectedSKUsAsDuplicates.Parameters.Add(outMsgParam);
                        conn.Open();
                        cmdUpdateSelectedSKUsAsDuplicates.ExecuteNonQuery();
                        strOutput = Convert.ToString(cmdUpdateSelectedSKUsAsDuplicates.Parameters["@Result"].Value);
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strOutput = "Error: " + ex.Message;
            }

            return strOutput;
        }
        #endregion

        #region Fetch Customer Input File SKUs updated in SQL table
        public DataTable FetchCIFSKUsFromTable(ConfirmDuplicateSKUsModel model)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spProjectDuplicateSKUsFromCIF", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", model.CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", model.ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", model.BatchNo);
                        cmd.Parameters.AddWithValue("@SelectedColumns", model.SelectedColumns);
                        cmd.Parameters.AddWithValue("@Mode", 2);
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