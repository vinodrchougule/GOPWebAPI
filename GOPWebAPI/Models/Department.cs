using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class Department
    {
        public int SlNo { get; set; }

        [Key]
        public int DepartmentID { get; set; }

        [Required]
        [StringLength(100)]
        public string Name { get; set; }

        [Required]
        public bool IsActive { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class DepartmentHeadCount
    {
        public string Name { get; set; }

        public int HeadCount { get; set; }
    }
}