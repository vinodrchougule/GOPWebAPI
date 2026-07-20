using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALAccessControl
    {
        private readonly string _connectionString;

        public DALAccessControl(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Send email to management when attempt is maded to export data to Excel
        public void SendEmailToManagementAboutExportOfData(string PageName, string UserID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSendEmailToManagement";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@PageName", PageName);
                    cmd.Parameters.AddWithValue("@Subject", "Attempt to Export the data from GOP application");
                    cmd.Parameters.AddWithValue("@UserID", UserID);

                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
        }
        #endregion
    }
}