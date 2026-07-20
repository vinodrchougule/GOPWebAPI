using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public enum DataMode
    {
        Create=1,
        Update=2,
        Delete=3,
        ReadAll=4,
        ReadOne=5
    }
}