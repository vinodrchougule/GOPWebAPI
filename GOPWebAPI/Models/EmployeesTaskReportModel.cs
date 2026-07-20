using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class EmployeesTaskReportModel
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string Department { get; set; }
        public string Manager { get; set; }
        public string Activity { get; set; }
        public long ProductionAllocatedCount { get; set; }
        public long ProductionCompletedCount { get; set; }
        public long QCAllocatedCount { get; set; }
        public long QCCompletedCount { get; set; }
        public decimal ProductionHoursWorked { get; set; }
        public decimal QCHoursWorked { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public int NoOfProjects { get; set; }
        public long InputCount { get; set; }
        public int ProductionTarget { get; set; }
        public int QCTarget { get; set; }
        public decimal ProductivityManDays { get; set; }
        public decimal AveragePerDay { get; set; }
        public DateTime? ProductionAllocatedOn { get; set; }
        public DateTime? QCAllocatedOn { get; set; }
        public DateTime? ProductionStartDate { get; set; }
        public DateTime? ProductionEndDate { get; set; }
        public DateTime? QCStartDate { get; set; }
        public DateTime? QCEndDate { get; set; }

        public List<EmployeeCodes> EmployeeCodes { get; set; }
        public string UserID { get; set; }
    }

    public class EmployeesTaskSummaryReportModel
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string EmployeeNameCode { get; set; }
        public string Department { get; set; }
        public string Activity { get; set; }
        public int Projects { get; set; }
        public int Activities { get; set; }
        public long ProductionAllocatedCount { get; set; }
        public long ProductionCompletedCount { get; set; }
        public long QCAllocatedCount { get; set; }
        public long QCCompletedCount { get; set; }
        public decimal HoursWorked { get; set; }
        public decimal ManDays { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string Status { get; set; }
        public decimal ProductionAllocatedManDays { get; set; }
        public decimal QCAllocatedManDays { get; set; }
        public List<EmployeeCodes> EmployeeCodes { get; set; }
        public List<StatusOptions> StatusOptions { get; set; }
        public string UserID { get; set; }
    }

    public class UnAllocatedResourceModel
    {
        public int SlNo { get; set; }
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string Department { get; set; }
        public string Manager { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string UserID { get; set; }
    }

    public class EmployeeCodes
    {
        public string EmpCode { get; set; }
    }

    public class StatusOptions
    {
        public string Status { get; set; }
    }
}