using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;

namespace GOPWebAPI.BLL
{
    public class BLLQC
    {
        private readonly string _connectionString;
        private readonly DALQC _DALQC;
        public BLLQC(string connectionString)
        {
            _connectionString = connectionString;
            _DALQC = new DALQC(_connectionString);
        }

        #region Fetch Moved To QC Customer Codes
        public DataTable ReadMovedToQCCustomerCodes()
        {
            return _DALQC.ReadMovedToQCCustomerCodes();
        }
        #endregion

        #region Fetch Moved To QC Project Codes of Customer
        public DataTable ReadMovedToQCProjectCodes(string CustomerCode)
        {
            return _DALQC.ReadMovedToQCProjectCodes(CustomerCode);
        }
        #endregion

        #region Fetch Moved To QC Batch Nos. of Project
        public DataTable ReadMovedToQCBatchNos(string CustomerCode, string ProjectCode)
        {
            return _DALQC.ReadMovedToQCBatchNos(CustomerCode, ProjectCode);
        }
        #endregion

        #region Fetch Moved To QC Noun Modifiers
        public DataTable ReadMovedToQCNounModifiers(string CustomerCode, string ProjectCode, string BatchNo)
        {
            return _DALQC.ReadMovedToQCNounModifiers(CustomerCode, ProjectCode, BatchNo);
        }
        #endregion

        #region Fetch Moved To QC Project Items
        public DataTable ReadMovedToQCProjectItems(string CustomerCode, string ProjectCode, string BatchNo, string Noun, string Modifier,int PageNo,int PageSize, bool IsToFetchALLDetails)
        {
            return _DALQC.ReadMovedToQCProjectItems(CustomerCode, ProjectCode, BatchNo, Noun, Modifier,PageNo,PageSize, IsToFetchALLDetails);
        }
        #endregion
    }
}