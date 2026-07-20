using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Xml;
using System.Xml.Serialization;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.Incident_Report_Models;

namespace GOPWebAPI.DAL
{
    public class DALIncidentRegister
    {
        private readonly string _connectionString;

        public DALIncidentRegister(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Register an Incident
        public string RegisterIncident(IncidentRegisterModel objVar)
        {
            string strOutput = string.Empty;

            try
            {
                #region Departments Affected
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<DepartmentsAffected>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, objVar.DepartmentsAffectedList);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentRegister", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 1);  // Register / Create an incident flag
                        command.Parameters.AddWithValue("@IncidentType", objVar.IncidentType);
                        command.Parameters.AddWithValue("@DepartmentResolvingIncident", objVar.DepartmentResolvingIncident);
                        command.Parameters.AddWithValue("@IncidentDescription", objVar.IncidentDescription);
                        command.Parameters.AddWithValue("@NameOfPersonReportingIncident", objVar.NameOfPersonReportingIncident);
                        command.Parameters.AddWithValue("@ContactNo", objVar.ContactNo);
                        command.Parameters.AddWithValue("@EmailID", objVar.EmailID);
                        command.Parameters.AddWithValue("@IncidentDate", objVar.IncidentDate.ToString("dd-MMM-yyyy"));
                        command.Parameters.AddWithValue("@IncidentTime", objVar.IncidentTime);
                        command.Parameters.AddWithValue("@IncidentLocation", objVar.IncidentLocation);
                        command.Parameters.AddWithValue("@InformationAffected", objVar.InformationAffected);
                        command.Parameters.AddWithValue("@EquipmentAffected", objVar.EquipmentAffected);
                        command.Parameters.AddWithValue("@NoOfPeopleAffected", objVar.NoOfPeopleAffected);
                        command.Parameters.AddWithValue("@ImpactOnBusiness", objVar.ImpactOnBusiness);
                        command.Parameters.AddWithValue("@Priority", objVar.Priority);
                        //Send Departments Affected as xml
                        command.Parameters.Add(new SqlParameter("@DepartmentsAffected", SqlDbType.Xml)
                        {
                            Value = sqlXml
                        });
                        command.Parameters.AddWithValue("@AssetIDs", objVar.AssetIDs);
                        command.Parameters.AddWithValue("@IsConfirmed", objVar.IsConfirmed);
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

        #region Read All Incidents
        public DataTable ReadAllIncidents()
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentRegister", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 4);    //flag to read all incidents
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

        #region Read Incident By Id
        public DataTable ReadIncidentById(long IncidentRegisterID)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentRegister", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 5);
                        cmd.Parameters.AddWithValue("@IncidentRegisterID", IncidentRegisterID);
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

        #region Read Departments Affected By Incident Id
        public DataTable ReadDepartmentsAffectedById(long incidentRegisterID)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentDepartmentsAffected", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@IncidentRegisterID", incidentRegisterID);
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

        #region Read Incidents Unique Search Values by Search Field
        public DataTable ReadIncidentsUniqueSearchValuesBySearchField(string SearchField)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentRegisterFieldUniqueValues", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SearchField", SearchField);
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

        #region Read Incidents By Search Field and Search Value
        public DataTable ReadIncidentsBySearchFieldAndSearchValue(string SearchField, string SearchValue)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentSearchBySearchFieldAndValue", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SearchField", SearchField);
                        cmd.Parameters.AddWithValue("@SearchValue", SearchValue);
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

        #region Read Incidents Count Summary with status and each incident type for selected year and / or all
        public DataTable ReadIncidentsCountSummaryByYear(string YearOfIncident = null)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentsCountSummaryByYear", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Year", YearOfIncident);
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

        #region Read Incidents By Year and Status
        public DataTable ReadIncidentsByIncidentYearAndStatus(string Department, string IncidentType,string Status, string YearOfIncident=null)
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentsByYearAndStatus", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        if(!string.IsNullOrEmpty(YearOfIncident))
                            cmd.Parameters.AddWithValue("@Year", YearOfIncident);
                        cmd.Parameters.AddWithValue("@Department", Department);
                        cmd.Parameters.AddWithValue("@IncidentType", IncidentType);
                        cmd.Parameters.AddWithValue("@Status", Status);
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

        #region Update an Incident
        public string UpdateIncident(IncidentRegisterModel objVar)
        {
            string strOutput = string.Empty;
            try
            {
                #region Departments Affected
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<DepartmentsAffected>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, objVar.DepartmentsAffectedList);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentRegister", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 2);  // Example flag for update operation
                        command.Parameters.AddWithValue("@IncidentRegisterID", objVar.IncidentRegisterID);
                        command.Parameters.AddWithValue("@IncidentType", objVar.IncidentType);
                        command.Parameters.AddWithValue("@DepartmentResolvingIncident", objVar.DepartmentResolvingIncident);
                        command.Parameters.AddWithValue("@IncidentDescription", objVar.IncidentDescription);
                        command.Parameters.AddWithValue("@NameOfPersonReportingIncident", objVar.NameOfPersonReportingIncident);
                        command.Parameters.AddWithValue("@ContactNo", objVar.ContactNo);
                        command.Parameters.AddWithValue("@EmailID", objVar.EmailID);
                        command.Parameters.AddWithValue("@IncidentDate", objVar.IncidentDate.ToString("dd-MMM-yyyy"));
                        command.Parameters.AddWithValue("@IncidentTime", objVar.IncidentTime);
                        command.Parameters.AddWithValue("@IncidentLocation", objVar.IncidentLocation);
                        command.Parameters.AddWithValue("@InformationAffected", objVar.InformationAffected);
                        command.Parameters.AddWithValue("@EquipmentAffected", objVar.EquipmentAffected);
                        command.Parameters.AddWithValue("@NoOfPeopleAffected", objVar.NoOfPeopleAffected);
                        command.Parameters.AddWithValue("@ImpactOnBusiness", objVar.ImpactOnBusiness);
                        command.Parameters.AddWithValue("@Priority", objVar.Priority);
                        //Send Departments Affected as xml
                        command.Parameters.Add(new SqlParameter("@DepartmentsAffected", SqlDbType.Xml)
                        {
                            Value = sqlXml
                        });
                        command.Parameters.AddWithValue("@AssetIDs", objVar.AssetIDs);
                        command.Parameters.AddWithValue("@IncidentStatus", objVar.IncidentStatus);
                        command.Parameters.AddWithValue("@IsConfirmed", objVar.IsConfirmed);
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

        #region Delete Incident
        public string DeleteIncident(IncidentRegisterModel objVar)
        {
            string strOutput = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentRegister", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        // Input parameters to the stored procedure
                        command.Parameters.AddWithValue("@Mode", 3);  // Example flag for delete operation
                        command.Parameters.AddWithValue("@IncidentRegisterID", objVar.IncidentRegisterID);
                        command.Parameters.AddWithValue("@DepartmentResolvingIncident", objVar.DepartmentResolvingIncident);
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

        #region Update Action on Incident
        public string UpdateActionOnIncident(IncidentRegisterModel objVar)
        {
            string strOutput = string.Empty;
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Define the command and specify the stored procedure to use
                    using (SqlCommand command = new SqlCommand("spIncidentAction", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // Add parameters to the stored procedure
                        command.Parameters.AddWithValue("@IncidentRegisterID", objVar.IncidentRegisterID);
                        command.Parameters.AddWithValue("@RootCause", objVar.RootCause);
                        command.Parameters.AddWithValue("@CorrectiveAction", objVar.CorrectiveAction);
                        command.Parameters.AddWithValue("@PreventiveAction", objVar.PreventiveAction);
                        command.Parameters.AddWithValue("@IncidentStatus", objVar.IncidentStatus);
                        command.Parameters.AddWithValue("@ActionCompletedByUserName", objVar.ActionCompletedByUserName);
                        command.Parameters.AddWithValue("@ActionCompletedOn", objVar.ActionCompletedOn);
                        command.Parameters.AddWithValue("@Remarks", objVar.Remarks);
                        command.Parameters.AddWithValue("@IsActionConfirmed", objVar.IsActionConfirmed);
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

        #region Read Incident Years
        public DataTable ReadIncidentYears()
        {
            DataTable dtDetails = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spIncidentYears", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
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

    }
}