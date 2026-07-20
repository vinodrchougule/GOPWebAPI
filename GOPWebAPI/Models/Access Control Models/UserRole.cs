using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class UserRoles
    {
        public UserRoles()
        {
            UserRolesList = new List<UserRole>();
        }
        
        [Key]
        public long UserRoleID { get; set; }

        public List<UserRole> UserRolesList { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class UserRole
    {
        public int SlNo { get; set; }

        [Required]
        [StringLength(50)]
        public string UserName { get; set; }

        [Required]
        [StringLength(50)]
        public string RoleName { get; set; }

        [Required]
        public bool IsActive { get; set; }
    }

    public class UserRolePageAccess
    {
        public string RoleName { get; set; }

        public string PageName { get; set; }

        public bool IsActive { get; set; }
    }
}