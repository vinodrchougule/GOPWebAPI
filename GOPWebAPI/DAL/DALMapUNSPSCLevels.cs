using GOPWebAPI.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.DAL
{
    public class DALMapUNSPSCLevels
    {

        #region Map UNSPSC Levels and return the output
        public DataTable MapUNSPSCLevels(string sqlTableName, string UNSPSCVersion)
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATAddUNSPSCLevels", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = sqlTableName;
                        cmd.Parameters.Add("@UNSPSCVersion", SqlDbType.VarChar).Value = UNSPSCVersion;
                        cmd.CommandTimeout = 0;
                        conn.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dt;
        }
        #endregion
    }
}