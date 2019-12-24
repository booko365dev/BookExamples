using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace EJON
{
    class Program
    {
        static void Main(string[] args)
        {
            //==> CSOM
            ClientContext spCtx = LoginCsom();
            Web rootWeb = spCtx.Web;
            spCtx.Load(rootWeb);
            spCtx.ExecuteQuery();
            Console.WriteLine(rootWeb.Created.ToShortDateString());

            //==> PnP Core
            ClientContext spPnpCtx = LoginPnPCore();
            Web rootWebPnp = spPnpCtx.Web;
            spPnpCtx.Load(rootWebPnp);
            spPnpCtx.ExecuteQuery();
            Console.WriteLine(rootWebPnp.Created.ToShortDateString());

            //==> PnP Core direct login
            LoginPnPCoreDirectly();

            //==> REST
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            // Simple GET request without body
            using (var client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/created";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }

            //// Full POST query with data in the body
            //using (var client = new SPHttpClient(webUri, userName, password))
            //{
            //    object myPayload = new
            //    {
            //        __metadata = new { type = "SP.List" },
            //        Title = "NewTestListRest",
            //        BaseTemplate = 100,
            //        Description = "Test NewListRest",
            //        AllowContentTypes = true,
            //        ContentTypesEnabled = true
            //    };
            //    string endpointUrl = webUri + "/_api/web/lists";
            //    var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
            //    Console.WriteLine(data);
            //}

            Console.ReadLine();
        }

        static ClientContext LoginCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spUrl"]);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }

        static ClientContext LoginPnPCore()
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }

        static void LoginPnPCoreDirectly()
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            using (ClientContext spCtx =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]))
            {
                Web rootWeb = spCtx.Web;
                spCtx.Load(rootWeb);
                spCtx.ExecuteQuery();
                Console.WriteLine(rootWeb.Created.ToShortDateString());
            }
        }
    }

    class SPHttpClientHandler : HttpClientHandler
    {
        public SPHttpClientHandler(Uri webUri, string userName, string password)
        {
            CookieContainer = GetAuthCookies(webUri, userName, password);
            FormatType = FormatType.JsonVerbose;
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,
                                                CancellationToken cancellationToken)
        {
            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            if (FormatType == FormatType.JsonVerbose)
            {
                request.Headers.Add("Accept", "application/json;odata=verbose");
            }
            return base.SendAsync(request, cancellationToken);
        }

        private static CookieContainer GetAuthCookies(Uri webUri,
                                                string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(userName, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(webUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(webUri, authCookie);
            return cookieContainer;
        }

        public FormatType FormatType { get; set; }
    }

    public enum FormatType
    {
        JsonVerbose,
        Xml
    }

    class SPHttpClient : HttpClient
    {
        public SPHttpClient(Uri webUri, string userName, string password) : base(
                                new SPHttpClientHandler(webUri, userName, password))
        {
            BaseAddress = webUri;
        }

        public object ExecuteJson(string requestUri, HttpMethod method,
                                    IDictionary<string, string> headers, object payload,
                                    bool GetBinaryResponse = false)
        {
            HttpResponseMessage response;
            switch (method.Method)
            {
                case "POST":
                    DefaultRequestHeaders.Add("X-RequestDigest", RequestFormDigest());
                    if (headers != null)
                    {
                        foreach (var header in headers)
                        {
                            DefaultRequestHeaders.Add(header.Key, header.Value);
                        }
                    }
                    if ((payload != null) && (payload.GetType().Name == "FileStream"))
                    {
                        StreamContent requestContent = new StreamContent((Stream)payload);
                        requestContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");
                        response = PostAsync(requestUri, requestContent).Result;
                    }
                    else
                    {
                        StringContent requestContent = new StringContent(
                                                    JsonConvert.SerializeObject(payload));
                        requestContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");
                        response = PostAsync(requestUri, requestContent).Result;
                    }
                    break;
                case "GET":
                    response = GetAsync(requestUri).Result;
                    break;
                default:
                    throw new NotSupportedException(string.Format(
                                        "Method {0} is not supported", method.Method));
            }

            response.EnsureSuccessStatusCode();

            if (GetBinaryResponse == true)
            {
                var responseContentStream = response.Content.ReadAsStreamAsync().Result;
                return responseContentStream;
            }
            else
            {
                var responseContent = response.Content.ReadAsStringAsync().Result;
                return String.IsNullOrEmpty(responseContent) ? new JObject() :
                                                        JObject.Parse(responseContent);
            }
        }

        public object ExecuteJson<T>(string requestUri, HttpMethod method, T payload,
                                        bool GetBinaryResponse = false)
        {
            return ExecuteJson(requestUri, method, null, payload, GetBinaryResponse);
        }

        public object ExecuteJson(string requestUri, bool GetBinaryResponse = false)
        {
            return ExecuteJson(requestUri, HttpMethod.Get, null, default(string),
                                                                    GetBinaryResponse);
        }

        public string RequestFormDigest()
        {
            var endpointUrl = string.Format("{0}/_api/contextinfo", BaseAddress);
            var result = this.PostAsync(endpointUrl, new StringContent(
                                                            string.Empty)).Result;
            result.EnsureSuccessStatusCode();
            var content = result.Content.ReadAsStringAsync().Result;
            var contentJson = JObject.Parse(content);
            return contentJson["d"]["GetContextWebInformation"][
                                                        "FormDigestValue"].ToString();
        }
    }
}

