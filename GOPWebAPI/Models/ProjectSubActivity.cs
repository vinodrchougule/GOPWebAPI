using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProjectSubActivity
    {
        public int SlNo { get; set; }

        [Key]
        public int ProjectSubActivityID { get; set; }

        [Required]
        [StringLength(50)]
        public string SubActivity { get; set; }

        [Required]
        public bool IsActive { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}