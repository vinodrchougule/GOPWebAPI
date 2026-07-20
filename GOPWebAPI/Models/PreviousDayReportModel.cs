using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class PreviousDayReportModel
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string Department { get; set; }
        public string ManagerName { get; set; }
        public string Activity { get; set; }
        public long ProductionCompletedCount { get; set; }
        public long QCCompletedCount { get; set; }
    }
}