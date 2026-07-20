using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Helpers
{
    public static class DBConnInfo
    {
        #region Get Connection String
        public static string ConnectionString()
        {
            return System.Configuration.ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
        }
        #endregion
    }
}