using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.Incident_Report_Models
{
    public class IncidentTypeModel
    {
        [Key]
        public int IncidentTypeID { get; set; }

        [Required]
        [StringLength(50)]
        public string IncidentType { get; set; }

        [Required]
        public bool IsActive { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}