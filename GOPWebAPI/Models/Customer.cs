using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class Customer
    {
        [Key]
        public int CustomerID { get; set; }

        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        public int NoOfProjects { get; set; }

        public long InputCount { get; set; }
        
        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
        
    }
}