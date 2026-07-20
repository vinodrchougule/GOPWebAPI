using System;
using context = System.Web.HttpContext;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace GOPWebAPI.Helpers
{
    public static class ExceptionLogging
    {
        #region Send Exception To DB
        public static void SendExceptionToDB(Exception ex)
        {
            using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
            {
                string exceptionURL = context.Current.Request.Url.ToString();
                SqlCommand cmd = new SqlCommand("spLogException", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ExceptionMsg", ex.Message.ToString());
                cmd.Parameters.AddWithValue("@ExceptionType", ex.GetType().Name.ToString());
                cmd.Parameters.AddWithValue("@ExceptionURL", exceptionURL);
                cmd.Parameters.AddWithValue("@ExceptionSource", ex.StackTrace.ToString());
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }
        #endregion
    }
}