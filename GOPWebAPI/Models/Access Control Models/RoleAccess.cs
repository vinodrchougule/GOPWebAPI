using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class RoleAccess
    {
        public RoleAccess()
        {
            RoleAccessList = new List<RolePageAccess>();
        }

        [Key]
        public long RoleAccessID { get; set; }

        public List<RolePageAccess> RoleAccessList { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class RolePageAccess
    {
        public int SlNo { get; set; }

        [Required]
        [StringLength(50)]
        public string RoleName { get; set; }

        [Required]
        [StringLength(50)]
        public string PageName { get; set; }

        [Required]
        public bool IsActive { get; set; }
    }
}