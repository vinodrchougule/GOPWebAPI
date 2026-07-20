using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class UNSPSCVersions
    {
        public string Version { get; set; }
    }

    public class UNSPSCSearchCriteriaModel
    {
        [Required]
        public string TableName { get; set; }

        public string Keyword1 { get; set; }

        public string Keyword2 { get; set; }

        public string ANDStatus { get; set; }

        public string ORStatus { get; set; }

        public string DoNotContainStatus { get; set; }

        public string UNSPSCCode1 { get; set; }

        public string UNSPSCCode2 { get; set; }

        public string UNSPSCCode3 { get; set; }

        public string UNSPSCCode4 { get; set; }

        [Required]
        public int PageNo { get; set; }

        [Required]
        public int PageSize { get; set; }
    }
        
    public class UNSPSCSearchResultModel
    {
        public int RowNum { get; set; }

        public int TotalCount { get; set; }

        public string Code { get; set; }

        public string Category { get; set; }
    }

    public class UNSPSCCategoryModel
    {
        public string SegmentCode { get; set; }
        public string Segment { get; set; }
        public string FamilyCode { get; set; }
        public string Family { get; set; }
        public string ClassCode { get; set; }
        public string Class { get; set; }
        public string CommodityCode { get; set; }
        public string Commodity { get; set; }
        public string CategoryDefinition { get; set; }
    }
}