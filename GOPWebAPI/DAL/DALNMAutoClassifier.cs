using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using GOPWebAPI.Helpers;

namespace GOPWebAPI.DAL
{
    public class DALNMAutoClassifier
    {
        private readonly string _connectionString;
        public DALNMAutoClassifier(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Process NM Auto Classifier 
        public DataTable ProcessNMAutoClassifier(string InputFileTableName, string AbbreviationFileTableName, string StdNounModifierFileTableName)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spAssignNounModifiers", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@InputFileTableName", InputFileTableName);
                        cmd.Parameters.AddWithValue("@AbbreviationFileTableName", AbbreviationFileTableName);
                        cmd.Parameters.AddWithValue("@StdNounModifierFileTableName", StdNounModifierFileTableName);

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