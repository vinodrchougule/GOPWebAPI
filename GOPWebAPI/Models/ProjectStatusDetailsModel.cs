using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectStatusDetailsModel
    {
        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [StringLength(5)]
        public string ProjectCode { get; set; }

        [StringLength(20)]
        public string ChangeStatusTo { get; set; }

        public DateTime? DeliveredDate { get; set; }

        public int DeliveredCount { get; set; }

        public DateTime? OnHoldDate { get; set; }

        [StringLength(500)]
        public string OnHoldReason { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}