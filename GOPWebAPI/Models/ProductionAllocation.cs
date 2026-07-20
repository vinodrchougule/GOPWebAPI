using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProductionAllocation
    {
        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }

        [StringLength(4)]
        public string BatchNo { get; set; }

        [Required]
        [StringLength(100)]
        public string AllocatedFileName { get; set; }

        [Required]
        [StringLength(100)]
        public string UniqueColumnName { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class CustomerCodeProjectCodeBatchNo
    {
        public string CustomerCode { get; set; }

        public string ProjectCode { get; set; }

        public string BatchNo { get; set; }

        public long ProjectInputCount { get; set; } 
    }

    public class ProjectBatchScopeInputCount
    {
        public string CustomerCode { get; set; }

        public string ProjectCode { get; set; }

        public string BatchNo { get; set; }

        public string Scope { get; set; }

        public long InputCount { get; set; }

        public int IsProjectAllocated { get; set; }

        public int ProductionCompletedCount { get; set; }

        public decimal ProductionCompletedPercentage { get; set; }

        public int QCCompletedCount { get; set; }

        public decimal QCCompletedPercentage { get; set; }

        public bool IsProjectSettingsExist { get; set; }
    }

    public class ProjectActivitiesCount
    {
        public int Index { get; set; }

        public long ProductionAllocationID { get; set; }

        public string Activity { get; set; }

        public int ActivityCount { get; set; }

        public int ProductionPendingToAllocate { get; set; }

        public int ProductionAllocatedCount { get; set; }
        
        public int ProductionCompletedCount { get; set; }

        public int ProductionPendingCount { get; set; }

        public decimal ProductionCompletionPercentage { get; set; }

        public int ProductionErrorCount { get; set; }

        public int IsAllocationDownloadedForProduction { get; set; }

        public int IsProductionErrorDownloaded { get; set; }
    }

    public class ProjectAllocationAndCompletedCountStatus
    {
        public int Index { get; set; }

        public string Activities { get; set; }

        public int ProductionAllocatedCount { get; set; }

        public int ProductionPendingToAllocate { get; set; }

        public int ProductionCompletedCount { get; set; }

        public decimal ProductionCompletionPercentage { get; set; }
    }

    public class UniqueColumn
    {
        public string  UniqueColumnName { get; set; }
    }

    public class ProductionExistingAllocation
    {
        public long ProductionAllocationID { get; set; }

        public DateTime AllocatedOn { get; set; }

        public string AllocatedByUserName { get; set; }

        public string AllocatedFileName { get; set; }

        public int AllocatedCount { get; set; }

        public int CompletedCount { get; set; }
    }

    public class ProductionAllocationDetails
    {
        public int ProductionAllocationDetailsID { get; set; }          //temporary added since bootstrap table requires a key to render table

        public long ProductionAllocationID { get; set; }

        public string Activities { get; set; }

        public string ProductionUser { get; set; }

        public string ChangeToProductionUser { get; set; }

        public string UserID { get; set; }

        public int ProductionAllocatedCount { get; set; }

        public int ProductionPendingCount { get; set; }

        public int ProductionCompletedCount { get; set; }
    }

    public class ProductionAllocationChangeSKUProdUser
    {
        public long ProductionAllocationID { get; set; }

        public string UniqueColumnValue { get; set; }

        public string Activities { get; set; }

        public string ChangeToProductionUser { get; set; }

        public string UserID { get; set; }
    }

    public class ProductionPendingSKU
    {
        public long ProductionAllocationID { get; set; }

        public string UniqueColumnName { get; set; }

        public string UniqueColumnNameValues { get; set; }

        public string Activity { get; set; }
    }

    public class OSPASKUs
    {
        [Required]
        public string CustomerCode { get; set; }
        
        [Required]
        public string ProjectCode { get; set; }

        public string BatchNo { get; set; }

        public List<CIFSKUID> CIFIDs { get; set; }

        [Required]
        public string AllocateToUser { get; set; }

        [Required]
        public string UserID { get; set; }
    }

    public class CIFSKUID
    {
        public int CIFID { get; set; }
    }

    public class OSPAAllocatedSKUs
    {
        public int ID { get; set; } 
        public string ShortDescription { get; set; }
        public string LongDescription { get; set; }
        public string UOM { get; set; }
        public string MFRName { get; set; }
        public string MFRPN { get; set; }
        public string VendorName { get; set; }
        public string VendorPN { get; set; }
        public string Status { get; set; }
        public string ProductionUser { get; set; }
    }


}