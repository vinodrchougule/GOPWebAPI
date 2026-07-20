using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Web;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models.GAT_Models;

namespace GOPWebAPI.DAL
{
    public class DALGAT
    {
        private readonly string _connectionString;

        public DALGAT(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Create sql Table in database
        public string CreateSQLTable(string spName)
        {
            string sqlTableName = string.Empty;

            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    using (SqlCommand cmdCreateSQLTable = new SqlCommand(spName, conn))
                    {
                        sqlTableName = "tmp" + Guid.NewGuid().ToString().Replace("-", "");
                        cmdCreateSQLTable.CommandType = CommandType.StoredProcedure;
                        cmdCreateSQLTable.Parameters.AddWithValue("@TableName", sqlTableName);
                        conn.Open();
                        cmdCreateSQLTable.CommandTimeout = 0;
                        cmdCreateSQLTable.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return sqlTableName;
        }
        #endregion

        #region Write File Data To SQL Server Table
        public bool WriteFileDataToSQLServerTable(string FileName,string sqlTableName, int WorksheetNo = 0)
        {
            bool IsSucceeded = false;

            try
            {
                string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    DataFormatConverter dataFormatConverter = new DataFormatConverter();
                    DataTable excelData = dataFormatConverter.ExcelToDataTable(UploadedFilepath, WorksheetNo);
                    SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                    bulkcopy.DestinationTableName = sqlTableName;
                    conn.Open();
                    bulkcopy.WriteToServer(excelData);
                    conn.Close();
                    IsSucceeded = true;
                }
            }
            catch (Exception ex)
            {
                if (sqlTableName != "tmp")
                {
                    DataFormatConverter dataFormatConverter1 = new DataFormatConverter();
                    dataFormatConverter1.DropTable(sqlTableName);
                }

                ExceptionLogging.SendExceptionToDB(ex);
            }

            return IsSucceeded;
        }
        #endregion

        #region Write List Data to SQL Server Table
        public bool WriteListDataToSQLServerTable(List<GATMapUNSPSC> UNSPSCDataList, string sqlTableName)
        {
            bool IsSucceeded = false;

            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    DataFormatConverter dataFormatConverter = new DataFormatConverter();
                    DataTable excelData = UNSPSCDataList.AsDataTable();
                    SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                    bulkcopy.DestinationTableName = sqlTableName;
                    conn.Open();
                    bulkcopy.WriteToServer(excelData);
                    conn.Close();
                    IsSucceeded = true;
                }
            }
            catch (Exception ex)
            {
                if (sqlTableName != "tmp")
                {
                    DataFormatConverter dataFormatConverter1 = new DataFormatConverter();
                    dataFormatConverter1.DropTable(sqlTableName);
                }

                ExceptionLogging.SendExceptionToDB(ex);
            }
            return IsSucceeded;
        }
        #endregion

        #region Transpose data from H to V and get it
        public DataTable TransposeDataFromHtoV(string sqlTableName)
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    using (SqlCommand cmd = new SqlCommand("spHtoVTransposeAndGetData", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = sqlTableName;
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

        #region Abbreviate Input File data and return the output
        public DataTable AbbreviateInputFile(string inputFileSqlTableName, string abbreviationFileSqlTableName)
        {
            DataTable dt = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATABBGetData", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@InputFileTableName", SqlDbType.VarChar).Value = inputFileSqlTableName;
                        cmd.Parameters.Add("@AbbreviationFileTableName", SqlDbType.VarChar).Value = abbreviationFileSqlTableName;
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