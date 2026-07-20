using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.Incident_Report_Models
{
    public class IncidentRegisterModel
    {
        [Key]
        public long IncidentRegisterID { get; set; }
        public string IncidentNo { get; set; }
        [Required]
        public DateTime IncidentDate { get; set; }
        [Required]
        public string IncidentTime { get; set; }
        public int DepartmentIDResolvingIncident { get; set; }
        [Required]
        [StringLength(100)]
        public string DepartmentResolvingIncident { get; set; }
        public int IncidentTypeID { get; set; }
        [Required]
        [StringLength(50)]
        public string IncidentType { get; set; }
        [Required]
        [StringLength(4000)]
        public string IncidentDescription { get; set; }
        [RegularExpression("^[PIC]$")]
        public string IncidentStatus { get; set; }
        [Required]
        [StringLength(100)]
        public string NameOfPersonReportingIncident { get; set; }
        [StringLength(50)]
        public string ContactNo { get; set; }
        [StringLength(50)]
        public string EmailID { get; set; }
        [Required]
        [StringLength(50)]
        public string IncidentLocation { get; set; }
        [StringLength(100)]
        public string InformationAffected { get; set; }
        [StringLength(100)]
        public string EquipmentAffected { get; set; }
        public int NoOfPeopleAffected { get; set; }
        [StringLength(1)]
        [RegularExpression("^[LMH]$")]
        public string ImpactOnBusiness { get; set; }
        [StringLength(1)]
        [RegularExpression("^[LMH]$")]
        public string Priority { get; set; }
        public List<DepartmentsAffected> DepartmentsAffectedList { get; set; }
        public string DepartmentsAffectedCSV { get; set; }

        [StringLength(1000)]
        public string AssetIDs { get; set; }
        public bool IsConfirmed { get; set; } = false;
        public string RootCause { get; set; }
        public string CorrectiveAction { get; set; }
        public string PreventiveAction { get; set; }
        public int ActionCompletedByUserID { get; set; }
        public string ActionCompletedByUserName { get; set; }
        public DateTime? ActionCompletedOn { get; set; }
        public string Remarks { get; set; }
        public bool IsActionConfirmed { get; set; }
        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class DepartmentsAffected
    {
        public string DepartmentName { get; set; }
    }

    public class IncidentCountSummary
    {
        public string Department { get; set; }
        public string IncidentType { get; set; }
        public int PendingCount { get; set; }
        public int InProgressCount { get; set; }
        public int CompletedCount { get; set; }
        public int TotalCount { get; set; }
    }

    public class SearchIncidentUniqueFieldValues
    {
        public string SearchValue { get; set; }
    }
}