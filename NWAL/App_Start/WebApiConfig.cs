using System.Web.Http;

namespace NWAL
{
    public static class WebApiConfig
    {
        //gavdcodebegin 06
        public static void Register(HttpConfiguration config)
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
        //gavdcodeend 06
    }
}
