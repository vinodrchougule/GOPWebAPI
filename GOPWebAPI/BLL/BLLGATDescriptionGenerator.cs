using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using GOPWebAPI.DAL;
using GOPWebAPI.Models.GAT_Models;

namespace GOPWebAPI.BLL
{
    public class BLLGATDescriptionGenerator
    {
        private readonly string _connectionString;
        private readonly DALGATDescriptionGenerator _DALGATDescriptionGenerator;
        public BLLGATDescriptionGenerator(string connectionString)
        {
            _connectionString = connectionString;
            _DALGATDescriptionGenerator = new DALGATDescriptionGenerator(_connectionString);
        }

        #region Read Description Generator Setting Names
        public DataTable ReadDGSettingNames()
        {
            return _DALGATDescriptionGenerator.ReadDGSettingNames();
        }
        #endregion

        #region Save Description Generator Settings
        public string SaveDescriptionGeneratorSettings(DescriptionGenerator model)
        {
            return _DALGATDescriptionGenerator.SaveDescriptionGeneratorSettings(model);
        }
        #endregion

        #region Read Description Generator Saved Setting
        public DataTable ReadDescriptionGeneratorSavedSetting(string SettingName)
        {
            return _DALGATDescriptionGenerator.ReadDescriptionGeneratorSavedSetting(SettingName);
        }
        #endregion
    }
}