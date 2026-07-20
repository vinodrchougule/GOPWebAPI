using GOPWebAPI.DAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.BLL
{
    public class BLLAccessControl
    {
        private readonly string _connectionString;
        private readonly DALAccessControl _DALAccessControl;

        public BLLAccessControl(string connectionString)
        {
            _connectionString = connectionString;
            _DALAccessControl = new DALAccessControl(_connectionString);
        }

        #region Send email to management when attempt is maded to export data to Excel
        public void SendEmailToManagementAboutExportOfData(string PageName, string UserID)
        {
            _DALAccessControl.SendEmailToManagementAboutExportOfData(PageName, UserID);
        }
        #endregion
    }
}