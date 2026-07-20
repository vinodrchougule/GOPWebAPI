using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class EmployeeSpecificSummaryReportModel
    {
        public int SlNo { get; set; }
        public DateTime ProductionOrQCCompletedOn { get; set; }
        public decimal ProductiveHours { get; set; }
    }

    public class EmployeeSpecificDetailsReportModel
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string Activity { get; set; }
        public DateTime ProductionOrQCCompletedOn { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int ProductionTarget { get; set; }
        public decimal ProductionProductiveHours { get; set; }
        public int QCCompletedCount { get; set; }
        public int QCTarget { get; set; }
        public decimal QCProductiveHours { get; set; }
    }
}