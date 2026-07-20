using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectStatusReportModel
    {
        public int SlNo { get; set; }
        public string  EmployeeCode { get; set; }
        public string  EmployeeName { get; set; }
        public string  Activity { get; set; }
        public int ProductionAllocatedCount { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int ProductionPendingCount { get; set; }
        public int QCAllocatedCount { get; set; }
        public int QCCompletedCount { get; set; }
        public int QCPendingCount { get; set; }
        public DateTime? LastUpdatedDate { get; set; }
    }

    public class ProjectDetailsModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string Scope { get; set; }
        public long InputCount { get; set; }
        public DateTime ReceivedOn { get; set; }
        public DateTime? DeliveredOn { get; set; }
    }

    public class ProjectStatusActivitySummaryModel
    {
        public int index { get; set; }
        public string Activity { get; set; }
        public int ActivityCount { get; set; }
        public int ProductionAllocatedCount { get; set; }
        public decimal ProductionAllocatedPercentage { get; set; }
        public int ProductionCompletedCount { get; set; }
        public decimal ProductionCompletedPercentage { get; set; }
        public int QCAllocatedCount { get; set; }
        public decimal QCAllocatedPercentage { get; set; }
        public int QCCompletedCount { get; set; }
        public decimal QCCompletedPercentage { get; set; }
    }
}