using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class NounModifierAttribute
    {
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string Attribute { get; set; }
        public string MandatoryOrOptional { get; set; }
    }

    public class NMAV
    {
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string Attribute { get; set; }
        public string Value { get; set; }
    }
}