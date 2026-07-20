using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class UserPageAccess
    {
        public string UserName { get; set; }

        public string PageName { get; set; }

        public bool canAccess { get; set; }
    }
}