using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Xml;
using System.Xml.Serialization;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;

namespace GOPWebAPI.DAL
{
    public class DALQC
    {
        private readonly string _connectionString;

        public DALQC(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Fetch Moved To QC Customer Codes
        public DataTable ReadMovedToQCCustomerCodes()
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMovedToQCCustomerAndProjectCodesAndBatchNos", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 1);    //flag to read customer codes
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion

        #region Fetch Moved To QC Project Codes of Customer
        public DataTable ReadMovedToQCProjectCodes(string CustomerCode)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMovedToQCCustomerAndProjectCodesAndBatchNos", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 2);    //flag to read project codes
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion

        #region Fetch Moved To QC Batch Nos. of Project
        public DataTable ReadMovedToQCBatchNos(string CustomerCode, string ProjectCode)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMovedToQCCustomerAndProjectCodesAndBatchNos", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@Mode", 3);    //flag to read batch nos.
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion

        #region Fetch Moved To QC Noun Modifiers
        public DataTable ReadMovedToQCNounModifiers(string CustomerCode, string ProjectCode, string BatchNo)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMovedToQCNounModifiersList", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion

        #region Fetch Moved To QC Project Items
        public DataTable ReadMovedToQCProjectItems(string CustomerCode, string ProjectCode,string BatchNo, string Noun, string Modifier, int PageNo, int PageSize, bool IsToFetchALLDetails)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spMovedToQCProjectItems", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                        cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                        cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                        cmd.Parameters.AddWithValue("@Noun", Noun);
                        cmd.Parameters.AddWithValue("@Modifier", Modifier);
                        cmd.Parameters.AddWithValue("@PageNo", PageNo);
                        cmd.Parameters.AddWithValue("@PageSize", PageSize);
                        cmd.Parameters.AddWithValue("@IsToFetchALLDetails", IsToFetchALLDetails);
                        cmd.Parameters.Add("@Result", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion
    }
}