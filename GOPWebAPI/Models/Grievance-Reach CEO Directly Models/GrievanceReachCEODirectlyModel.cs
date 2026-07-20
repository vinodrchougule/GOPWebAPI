using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.Grievance_Reach_CEO_Directly_Models
{
    public class GrievanceReachCEODirectlyModel
    {
        [Key]
        public long SuggestionToManagementID { get; set; }

        [Required]
        [StringLength(100)]
        public string Subject { get; set; }

        [Required]
        [StringLength(1000)]
        public string Details { get; set; }

        public DateTime? CreatedOn { get; set; }

    }
}