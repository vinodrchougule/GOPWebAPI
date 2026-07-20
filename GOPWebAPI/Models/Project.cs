using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class Project
    {
        public Project()
        {
            Activities = new List<ActivityTarget>();
        }

        public int SlNo { get; set; }

        [Key]
        public long ProjectID { get; set; }

        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [StringLength(5)]
        public string ProjectCode { get; set; }

        [Required]
        [StringLength(1)]
        public string ProjectType { get; set; }

        [Required]
        [StringLength(3)]
        public string LocationCode { get; set; }

        [Required]
        [StringLength(1)]
        public string TypeOfInput { get; set; }

        [Required]
        public int InputCount { get; set; }

        [Required]
        [StringLength(1)]
        public string InputCountType { get; set; }

        [Required]
        public DateTime ReceivedDate { get; set; }

        [Required]
        public string ReceivedFormat { get; set; }

        [Required]
        public string OutputFormat { get; set; }

        [Required]
        public DateTime PlannedStartDate { get; set; }

        [Required]
        [StringLength(1)]
        public string DeliveryMode { get; set; }

        public DateTime? PlannedDeliveryDate { get; set; }

        public string DeliveryPlanFileName { get; set; }

        public bool IsResourceBased { get; set; }

        public string Remarks { get; set; }

        [Required]
        public string CustomerInputFileName { get; set; }

        public string Scope { get; set; }
        public string ScopeFileName { get; set; }

        public string Guideline { get; set; }
        public string GuidelineFileName { get; set; }

        public string Checklist { get; set; }
        public string ChecklistFileName { get; set; }

        public DateTime? EmailDate { get; set; }

        [StringLength(4000)]
        public string EmailDescription { get; set; }

        public string UNSPSCVersion { get; set; }

        public string MRODictionaryVersion { get; set; }

        public string Department { get; set; }

        public string CreatedByEmployeeName { get; set; }

        public int NoOfBatches { get; set; }

        public int NoOfActivities { get; set; }

        public string Status { get; set; }

        public bool canUserChangeProjectStatus { get; set; }

        public DateTime? HoldOnDate { get; set; }

        public string HoldOnReason { get; set; }

        public DateTime? DeliveredOn { get; set; }

        public long DeliveredCount { get; set; }

        public long ProductionAllocatedCount { get; set; }

        public long ProductionCompletedCount { get; set; }

        public long ProductionPendingCount { get; set; }

        public long QCAllocatedCount { get; set; }

        public long QCCompletedCount { get; set; }

        public long QCPendingCount { get; set; }

        public int NoOfResources { get; set; }

        public DateTime? FromDate { get; set; }

        public DateTime? ToDate { get; set; }

        public bool IsProjectSettingsExist { get; set; }

        public List<ActivityTarget> Activities { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class ActivityTarget
    {
        public string Activity { get; set; }
        public int NoOfSKUs { get; set; }
        public int ProductionTarget { get; set; }
        public int QCTarget { get; set; }
        public int QATarget { get; set; }
        public int AllocatedCount { get; set; }
    }

    public class ProjectActivitiesWithHoursWorked
    {
        public int SlNo { get; set; }
        public int ProjectActivityID { get; set; }
        public string ProjectActivity { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCCompletedCount { get; set; }
        public int HoursWorked { get; set; }
    }

    public class ProjectActivityResourcesWithHoursWorked
    {
        public int SlNo { get; set; }
        public string UserID { get; set; }
        public string Username { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int ProductionCompletedCount { get; set; }
        public int QCCompletedCount { get; set; }
        public int HoursWorked { get; set; }
    }

    public class ProjectSettings
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string AdditionalInfoPrefix { get; set; }
        public string MFRNamePrefix { get; set; }
        public string MFRPNPrefix { get; set; }
        public string VendorNamePrefix { get; set; }
        public string VendorPNPrefix { get; set; }
        public bool IsToIncludeAdditionalInfoInShortDesc { get; set; }
        public bool IsToIncludeAdditionalInfoInLongDesc { get; set; }
        public bool IsToIncludeMFRNameInShortDesc { get; set; }
        public bool IsToIncludeMFRNameInLongDesc { get; set; }
        public bool IsToIncludeMFRPNInShortDesc { get; set; }
        public bool IsToIncludeMFRPNInLongDesc { get; set; }
        public bool IsToIncludeVendorNameInShortDesc { get; set; }
        public bool IsToIncludeVendorNameInLongDesc { get; set; }
        public bool IsToIncludeVendorPNInShortDesc { get; set; }
        public bool IsToIncludeVendorPNInLongDesc { get; set; }

        public bool IsToConvertAttributeValueToUppercase { get; set; }

        public string MFRNameInputColumnName { get; set; }
        public string MFRPNInputColumnName { get; set; }
        public string VendorNameInputColumnName { get; set; }
        public string VendorPNInputColumnName { get; set; }
        public string ShortDescriptionInputColumnName { get; set; }
        public string LongDescriptionInputColumnName { get; set; }
        public string CustomColumnName1 {  get; set; }
        public string CustomColumnName2 { get; set; }
        public string CustomColumnName3 { get; set; }

        public List<ProjectSettingsSpecialCharacters> SpecialCharacters { get; set; }

        public string UserID { get; set; }
    }

    public class ProjectSettingsSpecialCharacters
    {
        public string Characters { get; set; }
    }

    public class ProjectUpdateDetails
    {
        [Key]
        public long ProjectUpdateDetailsID { get; set; }

        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }

        [Required]
        [StringLength(100)]
        public string Subject { get; set; }

        [Required]
        [StringLength(4000)]
        public string Details { get; set; }

        [StringLength(100)]
        public string UserUploadedFileName { get; set; }
        
        public string UserUploadedTempFileName { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}