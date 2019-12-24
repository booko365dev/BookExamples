using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace ESSV
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginCsom();
            ClientContext spAdminCtx = LoginAdminCsom();

            //SpCsCsomGetPropertiesTenant(spAdminCtx);
            //SpCsCsomGetValuePropertyTenant(spAdminCtx);
            //SpCsCsomUpdateValuePropertyTenant(spAdminCtx);

            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            Uri webBaseUri = new Uri(ConfigurationManager.AppSettings["spBaseUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            //SpCsRestFindAppCatalog(webBaseUri, userName, password);
            //SpCsRestFindTenantProps(webBaseUri, userName, password);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void SpCsCsomGetPropertiesTenant(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);

            foreach (PropertyInfo oneProperty in myTenant.GetType().GetProperties())
            {
                Console.WriteLine(oneProperty.Name);
            }
        }

        static void SpCsCsomGetValuePropertyTenant(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);

            spAdminCtx.Load(myTenant);
            spAdminCtx.ExecuteQuery();

            bool myAccessDevices = myTenant.BlockAccessOnUnmanagedDevices;
            Console.WriteLine(myAccessDevices);
        }

        static void SpCsCsomUpdateValuePropertyTenant(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);

            myTenant.BlockAccessOnUnmanagedDevices = false;
            myTenant.Update();
            spAdminCtx.ExecuteQuery();
        }

        static void SpCsRestFindAppCatalog(Uri webBaseUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webBaseUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webBaseUri +
                                "/_api/SP_TenantSettings_Current";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }

        static void SpCsRestFindTenantProps(Uri webBaseUri, string userName,
                                                                    string password)
        {
            Uri catalogUri = new Uri(webBaseUri + "/sites/appcatalog");
            using (SPHttpClient client = new SPHttpClient(catalogUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = catalogUri +
                                "/_api/web/GetStorageEntity('SomeKey')";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }

        //-------------------------------------------------------------------------------
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

        static ClientContext LoginCsom(string WebFullUrl)
        {
            ClientContext rtnContext = new ClientContext(WebFullUrl);

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

        static ClientContext LoginAdminCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spAdminUrl"]);

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
    }

    //-----------------------------------------------------------------------------------
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
                        StringContent requestContent = null; // = new StringContent(string.Empty);
                        if (payload.GetType().FullName == "System.String")
                        {
                            requestContent = new StringContent((string)payload);
                        }
                        else
                        {
                            requestContent = new StringContent(
                                                   JsonConvert.SerializeObject(payload));
                        }
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

            //response.EnsureSuccessStatusCode();

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
