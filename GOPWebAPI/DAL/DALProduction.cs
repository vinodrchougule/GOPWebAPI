using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.DynamicData;
using System.Web.UI.WebControls;
using GOPWebAPI.Helpers;

namespace GOPWebAPI.DAL
{
    public class DALProduction
    {
        private readonly string _connectionString;
        public DALProduction(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Read Project Overall Status all results
        public DataSet ReadProjectOverallStatusAllResults(string CustomerCode, string ProjectCode, string BatchNo, string Status)
        {
            DataSet dataSet = new DataSet();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spProjectOverallStatus", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.AddWithValue("@Status", Status);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataSet);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataSet;
        }
        #endregion

        #region Read Date Range Status Report
        public DataTable ReadDateRangeStatus (DateTime FromDate, DateTime ToDate, string Status)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spDateRangeStatusReport", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);
                        cmd.Parameters.AddWithValue("@Status", Status);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion

        #region User Based Status Report
        public DataTable ReadProductionAndQCUserNames ()
        {
            DataTable dtUserNames = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spUserBasedStatusReport", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@flag", 1);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtUserNames);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dtUserNames;
        }

        public DataTable ReadUserBasedStatus(string UserName, string Status)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spUserBasedStatusReport", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@flag", 2);
                        cmd.Parameters.AddWithValue("@UserName", UserName);
                        cmd.Parameters.AddWithValue("@Status", Status);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion

        #region Project Level Quality Report
        #region Project Level Quality Report Count Stats
        public DataTable ReadProjectLevelQualityCountStats(string CustomerCode, string ProjectCode, string BatchNo)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spProjectLevelQualityReportCounts", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion

        #region Project Level Quality Report SKU Details
        public DataTable ReadProjectLevelQualityReportSKUDetails(string CustomerCode, string ProjectCode, string BatchNo, string Status)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spProjectLevelQualityReportSKUs", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.AddWithValue("@Status", Status);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion
        #endregion

        #region Resource Level Quality Report
        #region Resource Level Quality Report Count Stats
        public DataTable ReadResourceLevelQualityCountStats(string UserName, string CustomerCode, string ProjectCode, string BatchNo, DateTime? FromDate, DateTime? ToDate)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spResourceLevelQualityReportCounts", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@UserName", UserName);
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion
        #region Resource Level Quality Report SKU Details
        public DataTable ReadResourceLevelQualityReportSKUDetails(string UserName, string CustomerCode, string ProjectCode, string BatchNo,DateTime? FromDate,DateTime? ToDate, string Status)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spResourceLevelQualityReportSKUDetails", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@UserName", UserName);
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.AddWithValue("@FromDate", FromDate);
                        cmd.Parameters.AddWithValue("@ToDate", ToDate);
                        cmd.Parameters.AddWithValue("@Status", Status);

                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dataTable;
        }
        #endregion
        #endregion

        #region Read MRO Ref. DB Part Details by Part No.
        public DataTable ReadMRORefDBPartDetailsByPartNo(string PartNo)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMRORefDBMFRPartDetails", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@MFRPartNo", PartNo);
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dataTable);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dataTable;
        }
        #endregion
    }
}