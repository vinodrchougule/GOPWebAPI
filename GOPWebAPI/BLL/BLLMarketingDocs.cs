using GOPWebAPI.DAL;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLMarketingDocs
    {
        private readonly string _connectionString;
        private readonly DALMarketingDocs _DALMarketingDocs;
        public BLLMarketingDocs(string connectionString)
        {
            _connectionString = connectionString;
            _DALMarketingDocs = new DALMarketingDocs(_connectionString);
        }

        #region Upload Marketing Document
        public string UploadMarketingDocument(MarketingDocModel model)
        {
            return _DALMarketingDocs.UploadMarketingDocument(model);
        }
        #endregion

        #region Read All Marketing Documents from input domain and doc type
        public DataTable ReadAllMarketingDocuments(string Domain, string DocType)
        {
            return _DALMarketingDocs.ReadAllMarketingDocuments(Domain, DocType);
        }
        #endregion

        #region Read Marketing Document Details by id
        public DataTable ReadMarketingDocumentById(int id)
        {
            return _DALMarketingDocs.ReadMarketingDocumentById(id);
        }
        #endregion

        #region Delete Marketing Document
        public string DeleteMarketingDocument(int id, string UserID)
        {
            return _DALMarketingDocs.DeleteMarketingDocument(id, UserID);
        }
        #endregion
    }
}