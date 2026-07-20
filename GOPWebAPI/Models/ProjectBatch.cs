using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectBatch
    {
        public int SlNo { get; set; }

        [Key]
        public long ProjectBatchID { get; set; }

        [Required]
        public long ProjectID { get; set; }

        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }

        public string BatchNo { get; set; }

        public string Scope { get; set; }

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
        public DateTime PlannedDeliveryDate { get; set; }

        public string Remarks { get; set; }

        [Required]
        public string CustomerInputFileName { get; set; }

        public string Status { get; set; }

        public DateTime? DeliveredDate { get; set; }

        public int DeliveredCount { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }

        public bool canUserDeliverProjectBatch { get; set; }

        public bool IsPostProjectBatchDetailsExist { get; set; }
    }

    public class ProjectBatchStatus
    {
        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }

        [Required]
        [StringLength(4)]
        public string BatchNo { get; set; }

        [Required]
        public DateTime DeliveredDate { get; set; }

        [Required]
        public int DeliveredCount { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}