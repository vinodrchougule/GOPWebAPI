using GOPWebAPI.Controllers;
using System.Web;
using System.Web.Mvc;

namespace GOPWebAPI
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
