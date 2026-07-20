using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class Account
    {
        public int SlNo { get; set; }

        [Key]
        public int UserID { get; set; }

        [Required]
        [StringLength(50)]
        public string FirstName { get; set; }

        [StringLength(50)]
        public string MiddleName { get; set; }

        [StringLength(50)]
        public string LastName { get; set; }

        [Required]
        [StringLength(50),MinLength(3)]
        public string UserName { get; set; }

        [StringLength(50), MinLength(6)]
        public string Password { get; set; }

        [Required]
        [StringLength(50)]
        public string Email { get; set; }

        [Required]
        [StringLength(100)]
        public string DepartmentName { get; set; }

        [Required]
        [StringLength(200)]
        public string ManagerName { get; set; }

        [Required]
        public bool IsLockedOut { get; set; }

        public DateTime? RelievingDate { get; set; }

        public string PhotoFileName { get; set; }

        public string PhotoFileBase64String { get; set; }

        [Required]
        [StringLength(50)]
        public string User { get; set; }
    }
}