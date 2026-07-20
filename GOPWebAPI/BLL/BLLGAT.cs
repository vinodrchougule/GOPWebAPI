using GOPWebAPI.DAL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models.GAT_Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLGAT
    {
        private readonly string _connectionString;
        private readonly DALGAT _DALGAT;
        public BLLGAT(string connectionString)
        {
            _connectionString = connectionString;
            _DALGAT = new DALGAT(_connectionString);
        }

        #region Create sql Table in database
        public string CreateSQLTable(string spName)
        {
            return _DALGAT.CreateSQLTable(spName);
        }
        #endregion

        #region Write File Data To SQL Server Table
        public bool WriteFileDataToSQLServerTable(string FileName, string sqlTableName, int WorksheetNo = 0)
        {
            return _DALGAT.WriteFileDataToSQLServerTable(FileName, sqlTableName, WorksheetNo);  
        }
        #endregion

        #region Write List Data to SQL Server Table
        public bool WriteListDataToSQLServerTable(List<GATMapUNSPSC> UNSPSCDataList, string sqlTableName)
        {
            return _DALGAT.WriteListDataToSQLServerTable(UNSPSCDataList, sqlTableName);
        }
        #endregion

        #region Transpose data from H to V and get it
        public DataTable TransposeDataFromHtoV(string sqlTableName)
        {
            return _DALGAT.TransposeDataFromHtoV(sqlTableName);
        }
        #endregion

        #region Abbreviate Input File data and return the output
        public DataTable AbbreviateInputFile(string inputFileSqlTableName, string abbreviationFileSqlTableName)
        {
            return _DALGAT.AbbreviateInputFile(inputFileSqlTableName, abbreviationFileSqlTableName);
        }
        #endregion

    }
}