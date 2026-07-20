using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class Dashboard
    {
        public int NoOfCustomers { get; set; }
        public int NoOfProjects { get; set; }
        public int NoOfOnGoingProjects { get; set; }
        public int NoOfCompletedProjects { get; set; }
        public int NoOfPendingProjects { get; set; }
        public decimal NoOfCompletedProjectsPercentage { get; set; }
        public decimal NoOfPendingProjectsPercentage { get; set; }
        public int NoOfBatches { get; set; }
        public int NoOfCompletedBatches { get; set; }
        public int NoOfPendingBatches { get; set; }
        public decimal NoOfCompletedBatchesPercentage { get; set; }
        public decimal NoOfPendingBatchesPercentage { get; set; }
        public int NoOfActiveTasks { get; set; }
        public int NoOfActiveResources { get; set; }
    }

    public class ActiveTasksDetails
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public int InputCount { get; set; }
        public string Activity { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCCompletedCount { get; set; }
    }

    public class ResourceDetails
    {
        public int SlNo { get; set; }
        public string ResourceCode { get; set; }
        public string ResourceName { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCCompletedCount { get; set; }
        public int HoursWorked { get; set; }
    }

    public class ResourceProductivityDetails
    {
        public int SlNo { get; set; }
        public DateTime DateWorked { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string Activity { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCCompletedCount { get; set; }
        public int ProductionTarget { get; set; }
        public int QCTarget { get; set; }
        public int HoursWorked { get; set; }
    }

    public class HoursWorkedDetails
    {
        public int SlNo { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string ResourceCode { get; set; }
        public string ResourceName { get; set; }
        public int NoOfHoursWorked { get; set; }
    }

    public class ProjectsCompletionStatus
    {
        public string Status { get; set; }
        public int NoOfProjects { get; set; }
    }
}