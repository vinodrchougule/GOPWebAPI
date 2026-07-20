using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace GOPWebAPI.Models
{
    public class Login
    {
        [Required]
        [StringLength(50), MinLength(3)]
        public string UserName { get; set; }

        [Required]
        [StringLength(50), MinLength(6)]
        public string Password { get; set; }
    }
}