using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class ItemSpendAnalysisModels
    {

    }

    class clsCommodity
    {
        public string Department { get; set; }
        public string Commodity { get; set; }
        public decimal HospitalVolume { get; set; }
        public decimal Total { get; set; }
        public decimal sumOfCommodityTotal { get; set; }
        public decimal EightyPercentOfCommodityTotal { get; set; }
    }

    class MBInput
    {
        public string SCACode { get; set; }
        public string GPO { get; set; }
        public string HealthSystemName { get; set; }
        public string FacilityName { get; set; }
        public string POC { get; set; }
        public string DeptArea { get; set; }
        public string CommodityTitle { get; set; }
        public string ItemNo { get; set; }
        public string Description { get; set; }
        public string Vendor { get; set; }
        public string VendorCat { get; set; }
        public string MFRName { get; set; }
        public string MFRCat { get; set; }
        public string UOM { get; set; }
        public string PKG { get; set; }
        public string HospitalQty { get; set; }
        public string HospitalPrice { get; set; }
        public string HospitalVolume { get; set; }
        public decimal decimalHospitalVolume
        {
            get
            {
                return Convert.ToDecimal(HospitalVolume.Replace("$", "").Replace("-", "0"));
            }
        }
    }
}