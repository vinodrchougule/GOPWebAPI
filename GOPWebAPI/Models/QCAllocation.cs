using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class QCAllocation
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
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class QCExistingAllocation
    {
        public long QCAllocationID { get; set; }

        public DateTime AllocatedOn { get; set; }

        public string AllocatedByUserName { get; set; }

        public string AllocatedFileName { get; set; }

        public int AllocatedCount { get; set; }

        public int CompletedCount { get; set; }
    }

    public class QCAllocationDetails
    {
        public int QCAllocationDetailsID { get; set; }          //temporary added since bootstrap table requires a key to render table

        public long QCAllocationID { get; set; }

        public string Activities { get; set; }

        public string QCUser { get; set; }

        public string ChangeToQCUser { get; set; }

        public string UserID { get; set; }

        public int QCAllocatedCount { get; set; }

        public int QCPendingCount { get; set; }

        public int QCCompletedCount { get; set; }
    }
}