using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectsSummaryReportModel
    {
        public int index { get; set; }
        public long ProjectID { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string Scope { get; set; }
        public long InputCount { get; set; }
        public DataTable ActivityDetails { get; set; }
    }
}