using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using ExcelDataReader;
using System.Text;

namespace GOPWebAPI.Helpers
{
    public class DataFormatConverter
    {
        public DataTable ExcelToDataTable(string storePath)
        {
            FileStream stream = File.Open(storePath, FileMode.Open, FileAccess.Read);

            string fileExtension = Path.GetExtension(storePath);
            IExcelDataReader excelReader = null;

            excelReader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            var dataSet = excelReader.AsDataSet(conf);

            excelReader.Close();
            stream.Close();
            return dataSet.Tables[0];
        }

        public DataTable ExcelToDataTable(string storePath, int WorkSheetNo)
        {
            FileStream stream = File.Open(storePath, FileMode.Open, FileAccess.Read);

            string fileExtension = Path.GetExtension(storePath);
            IExcelDataReader excelReader = null;

            excelReader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            var dataSet = excelReader.AsDataSet(conf);

            excelReader.Close();
            stream.Close();
            return dataSet.Tables[WorkSheetNo];
        }

        public string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        public void DropTable(string tableName)
        {
            SqlConnection conn1 = new SqlConnection(DBConnInfo.ConnectionString());

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn1;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "spDropTempTable";
            cmd.Parameters.AddWithValue("@TempTableName", tableName);
            cmd.CommandTimeout = 0;
            conn1.Open();
            cmd.ExecuteNonQuery();
            conn1.Close();
        }
    }
}