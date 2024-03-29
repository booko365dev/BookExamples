﻿using Microsoft.SharePoint.Client;
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
            // ATTENTION: Next routines using the deprecated SharePointPnPCoreOnline module
            //SpCsCsomExample();                //==> CSOM
            //SpCsRestExample01();              // Simple REST GET request without body
            //SpCsRestExample02();              // Full REST POST query with data in the body

            // ATTENTION: Next routines using the deprecated SharePointPnPCoreOnline module
            //SpCsPnPCoreExample();             //==> PnP Core  //*** LEGACY CODE ***
            //LoginPnPCoreDirectly();           //==> PnP Core direct login  //*** LEGACY CODE ***

            Console.ReadLine();
        }

        //-------------------------------------------------------------------------------

        //gavdcodebegin 007
        static void SpCsCsomExample()  //*** LEGACY CODE ***
        {
            ClientContext spCtx = LoginCsom();

            Web rootWeb = spCtx.Web;
            spCtx.Load(rootWeb);
            spCtx.ExecuteQuery();

            Console.WriteLine(rootWeb.Created.ToShortDateString());
        }
        //gavdcodeend 007

        //gavdcodebegin 008
        static void SpCsPnPCoreExample()  //*** LEGACY CODE ***
        {
            // ATTENTION: Using the deprecated SharePointPnPCoreOnline module
            ClientContext spPnpCtx = LoginPnPCore();

            Web rootWebPnp = spPnpCtx.Web;
            spPnpCtx.Load(rootWebPnp);
            spPnpCtx.ExecuteQuery();

            Console.WriteLine(rootWebPnp.Created.ToShortDateString());
        }
        //gavdcodeend 008

        //gavdcodebegin 009
        static void SpCsRestExample01()  //*** LEGACY CODE ***
        {
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            using (var client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/created";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);

                Console.WriteLine(data);
            }
        }
        //gavdcodeend 009

        //gavdcodebegin 010
        static void SpCsRestExample02()  //*** LEGACY CODE ***
        {
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            using (var client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.List" },
                    Title = "NewTestListRest",
                    BaseTemplate = 100,
                    Description = "Test NewListRest",
                    AllowContentTypes = true,
                    ContentTypesEnabled = true
                };
                string endpointUrl = webUri + "/_api/web/lists";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);

                Console.WriteLine(data);
            }
        }
        //gavdcodeend 010

        //-------------------------------------------------------------------------------

        //gavdcodebegin 001
        static ClientContext LoginCsom()  //*** LEGACY CODE ***
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
        //gavdcodeend 001

        //gavdcodebegin 002
        static ClientContext LoginPnPCore()  //*** LEGACY CODE ***
        {
            // ATTENTION: Using the deprecated SharePointPnPCoreOnline module
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }
        //gavdcodeend 002

        //gavdcodebegin 006
        static void LoginPnPCoreMFA()  //*** LEGACY CODE ***
        {
            // ATTENTION: Using the deprecated SharePointPnPCoreOnline module
            string siteUrl = ConfigurationManager.AppSettings["spUrl"];
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext mfaContext = authManager.GetWebLoginClientContext(siteUrl);
            
            Web myWeb = mfaContext.Web;
            mfaContext.Load(myWeb, wb => wb.Title);
            mfaContext.ExecuteQuery();
            
            Console.WriteLine("Connected to the site " + myWeb.Title + " using MFA");
        }
        //gavdcodeend 006

        //gavdcodebegin 003
        static void LoginPnPCoreDirectly()  //*** LEGACY CODE ***
        {
            // ATTENTION: Using the deprecated SharePointPnPCoreOnline module
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
        //gavdcodeend 003
    }

    //gavdcodebegin 004
    class SPHttpClientHandler : HttpClientHandler  //*** LEGACY CODE ***
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
    //gavdcodeend 004

    //gavdcodebegin 005
    class SPHttpClient : HttpClient  //*** LEGACY CODE ***
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
    //gavdcodeend 005
}
