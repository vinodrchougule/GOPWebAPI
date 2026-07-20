using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class OEMDescriptionGeneratorModels
    {
        
    }

    public class AdditionalFeatures
    {
        public string ReferenceID { get; set; }
        public string NM { get; set; }
        public string ClassName { get; set; }
        public string VendorCode { get; set; }
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public int Seq { get; set; }
        public string Flag { get; set; }
    }
}