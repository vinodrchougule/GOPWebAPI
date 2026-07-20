using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectAllocation
    {
        [Key]
        public long ProjectAllocationID { get; set; }
        
        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }
        
        public string BatchNo { get; set; }

        public string FileName { get; set; }

        public string UniqueColumnName { get; set; }

        public string AllocatedBy { get; set; }

        public DateTime? AllocatedOn { get; set; }
        
        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class EmployeeAllocation
    {
        public string EmployeeCode { get; set; }

        public int ProductionAllocated { get; set; }

        public int QCAllocated { get; set; }

        public int QAAllocated { get; set; }
    }
}