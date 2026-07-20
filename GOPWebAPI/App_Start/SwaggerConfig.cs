using System.Web.Http;
using WebActivatorEx;
using GOPWebAPI;
using Swashbuckle.Application;

[assembly: PreApplicationStartMethod(typeof(SwaggerConfig), "Register")]

namespace GOPWebAPI
{
    public class SwaggerConfig
    {
        public static void Register()
        {
            var thisAssembly = typeof(SwaggerConfig).Assembly;
            GlobalConfiguration.Configuration
              .EnableSwagger(c => c.SingleApiVersion("v1", "GOP WEB API"))
              .EnableSwaggerUi(); 
        }
    }
}