using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Http.ExceptionHandling;
using GOPWebAPI.Models.Incident_Report_Models;
using GOPWebAPI.Helpers;

namespace GOPWebAPI.DAL
{
    public class DALIncidentType
    {
        private readonly string _connectionString;

        public DALIncidentType(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Create Incident Type
        public string AddIncidentType(IncidentTypeModel objVar)
        {
            string strOutput = string.Empty;

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentType", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 1);  // Example flag for insert operation
                        command.Parameters.AddWithValue("@IncidentType", objVar.IncidentType);
                        command.Parameters.AddWithValue("@UserID", objVar.UserID);

                        // Output parameter to capture the message
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);

                        // Execute the stored procedure
                        command.ExecuteNonQuery();

                        // Retrieve the output message from the stored procedure
                        strOutput = Convert.ToString(command.Parameters["@Result"].Value);

                        connection.Close();
                    }

                    return strOutput;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strOutput = "Error: " + ex.Message;
                return strOutput;
            }
        }
        #endregion

        #region Read All Incident Types
        public DataTable ReadAllIncidentTypes()
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentType", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 4);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
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

        #region Read Incident Type By Id
        public DataTable ReadIncidentTypeById(int IncidentTypeID)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentType", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 5);
                        cmd.Parameters.AddWithValue("@IncidentTypeID", IncidentTypeID);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
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

        #region Update Incident Type
        public string UpdateIncidentType(IncidentTypeModel objVar)
        {
            string strOutput = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentType", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 2);  // Example flag for update operation
                        command.Parameters.AddWithValue("@IncidentTypeID", objVar.IncidentTypeID);
                        command.Parameters.AddWithValue("@IncidentType", objVar.IncidentType);
                        command.Parameters.AddWithValue("@IsActive", objVar.IsActive);
                        command.Parameters.AddWithValue("@UserID", objVar.UserID);

                        // Output parameter to capture the message
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);

                        // Execute the stored procedure
                        command.ExecuteNonQuery();

                        // Retrieve the output message from the stored procedure
                        strOutput = Convert.ToString(command.Parameters["@Result"].Value);

                        connection.Close();
                    }

                    return strOutput;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strOutput = "Error: " + ex.Message;
                return strOutput;  // Return false indicating an error occurred
            }
        }
        #endregion

        #region Delete Incident Type
        public string DeleteIncidentType(IncidentTypeModel objVar)
        {
            string strOutput = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentType", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 3);  // Example flag for delete operation
                        command.Parameters.AddWithValue("@IncidentTypeID", objVar.IncidentTypeID);
                        command.Parameters.AddWithValue("@UserID", objVar.UserID);

                        // Output parameter to capture the message
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);

                        // Execute the stored procedure
                        command.ExecuteNonQuery();

                        // Retrieve the output message from the stored procedure
                        strOutput = Convert.ToString(command.Parameters["@Result"].Value);

                        connection.Close();
                    }

                    return strOutput;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strOutput = "Error: " + ex.Message;
                return strOutput;  // Return false indicating an error occurred
            }
        }
        #endregion
    }
}