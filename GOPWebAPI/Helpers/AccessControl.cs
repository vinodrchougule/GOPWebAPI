using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;
using GOPWebAPI.Models;

namespace GOPWebAPI.Helpers
{
    public static class AccessControl
    {
        #region Validate User Page Access
        public static bool CanUserAccessPage(string UserName, string PageName)
        {
            System.Data.Common.DbDataReader sqlReader;

            using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
            {
                SqlCommand cmd = new SqlCommand("spValidateUserPageAccess", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", UserName);
                cmd.Parameters.AddWithValue("@PageName", PageName);
                
                conn.Open();
                UserPageAccess userPageAccess = new UserPageAccess();
                sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                while (sqlReader.Read())
                {
                    userPageAccess.UserName = sqlReader["UserName"].ToString();
                    userPageAccess.PageName = sqlReader["PageName"].ToString();
                    userPageAccess.canAccess = Convert.ToBoolean(sqlReader["canAccess"]);
                }
                conn.Close();

                if (userPageAccess.canAccess)
                    return true;
                
                return false;
            }
        }
        #endregion
    }
}