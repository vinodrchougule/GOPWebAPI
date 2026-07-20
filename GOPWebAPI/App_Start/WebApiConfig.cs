using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Owin;
using Newtonsoft.Json.Serialization;
using System.Web.Http.Cors;

namespace GOPWebAPI
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services
            config.Formatters.XmlFormatter.SupportedMediaTypes.Add(new System.Net.Http.Headers.MediaTypeHeaderValue("multipart/form-data"));
            //var enableCorsAttribute = new EnableCorsAttribute("*", "*", "*");
            //config.EnableCors(enableCorsAttribute);
            // Web API configuration and services
            // Configure Web API to use only bearer token authentication.
            //config.SuppressDefaultHostAuthentication();
            //config.Filters.Add(new HostAuthenticationFilter(OAuthDefaults.AuthenticationType));

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            config.Routes.MapHttpRoute("CustomApi", "api/{controller}/{action}/{id}", new { id = RouteParameter.Optional });

            config.Routes.MapHttpRoute(
               name: "DefaultApiWithAction",
               routeTemplate: "api/{controller}/{action}/{customercode}/{projectcode}/{id}",
               defaults: new { customercode = RouteParameter.Optional, projectcode = RouteParameter.Optional, id = RouteParameter.Optional }
           );
        }
    }
}
