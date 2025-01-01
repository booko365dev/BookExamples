using System.Web.Http;

namespace NWAL
{
    public static class WebApiConfig
    {
        //gavdcodebegin 006
        public static void Register(HttpConfiguration config)  // Legacy code
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            config.InitializeReceiveGenericJsonWebHooks();
        }
        //gavdcodeend 006
    }
}
