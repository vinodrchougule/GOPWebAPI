using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectActivity
    {
        public int SlNo { get; set; }

        [Key]
        public int ProjectActivityID { get; set; }

        [Required]
        [StringLength(50)]
        public string Activity { get; set; }

        [Required]
        public bool IsActive { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}