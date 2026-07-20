using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System.Data.SqlClient;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/AccessControl")]
    public class AccessControlController : ApiController
    {
        #region Can User Access Page
        [HttpGet]
        [Route("CanUserAccessPage/{UserName}/{PageName}")]
        public bool CanUserAccessPage(string UserName, string PageName)
        {
            if (AccessControl.CanUserAccessPage(UserName, PageName))
                return true;

            return false;
        }
        #endregion

        #region Read User Menu Access List
        [HttpGet]
        [Route("ReadUserMenuAccessList/{UserName}/{MenuName}")]
        public IHttpActionResult ReadUserMenuAccessList(string UserName,string MenuName)
        {
            List<UserPageAccess> UserPageAccessList = new List<UserPageAccess>();
            System.Data.Common.DbDataReader sqlReader;
            using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
            {
                //Initialize command object
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = "spMenuAccessControl";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;

                //Add parameters with values
                cmd.Parameters.AddWithValue("@UserName", UserName);
                cmd.Parameters.AddWithValue("@MenuName", MenuName);
                
                //Calling sp to get list of Menu Access
                conn.Open();
                sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                while (sqlReader.Read())
                {
                    UserPageAccess userPageAccess = new UserPageAccess();
                    userPageAccess.UserName = sqlReader["UserName"].ToString();
                    userPageAccess.PageName = sqlReader["PageName"].ToString();
                    userPageAccess.canAccess = Convert.ToBoolean(sqlReader["canAccess"]);
                    UserPageAccessList.Add(userPageAccess);
                }
                conn.Close();

                //return list to the request
                return Ok(UserPageAccessList);
            }
        }
        #endregion
    }
}
