using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class PeriodicProjectReportModel
    {
        public int SlNo { get; set; }
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string Activity { get; set; }
        public int ProductionAllocatedCount { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCAllocatedCount { get; set; }
        public int QCCompletedCount { get; set; }
    }

    public class PeriodicProjectReportInputModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string UserID { get; set; }
    }
}