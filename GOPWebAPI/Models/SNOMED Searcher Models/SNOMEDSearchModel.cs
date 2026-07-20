using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GOPWebAPI.Models
{
    public class SNOMEDSearchModel
    {
        public string SearchText1 { get; set; }
        public string SearchText2 { get; set; }
        public string SearchText3 { get; set; }
        public string SearchText4 { get; set;}
    }
    
    public class SNOMEDConceptTerm
    {
        public string Active { get; set;}
        public string IsFSN { get; set; }
        public string ConceptId { get; set;}
        public string Term { get; set;}
    }

    public class SNOMEDConceptSynonym
    {
        public string ConceptId { get; set; }
        public string Term { get; set; }
    }

    public class SNOMEDConceptParent
    {
        public string DestinationID { get; set; }
        public string DestinationIDName { get; set; }
    }

    public class SNOMEDConceptChildren
    {
        public string SourceId { get; set; }
        public string SourceIDName { get; set;}
    }
}
