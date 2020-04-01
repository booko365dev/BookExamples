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

namespace FCJY
{
    class Program
    {
        static void Main(string[] args)
        {
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            Uri webBaseUri = new Uri(ConfigurationManager.AppSettings["spBaseUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            //SpCsRestCreateOneCommunicationSiteCollection(webBaseUri, userName, password);
            //SpCsRestCreateOneSiteCollection(webBaseUri, userName, password);
            //SpCsRestCreateOneWebInSiteCollection(webUri, userName, password);
            //SpCsRestReadAllSiteCollections(webBaseUri, userName, password);
            //SpCsRestReadAllWebsInSiteCollection(webUri, userName, password);
            //SpCsRestUpdateOneWeb(webUri, userName, password);
            //SpCsRestDeleteOneWebFromSiteCollection(webUri, userName, password);
            //SpCsRestGetRoleDefinitionsWeb(webUri, userName, password);
            //SpCsRestFindUserPermissionsWeb(webUri, userName, password);
            //SpCsRestFindOtherUserPermissionsWeb(webUri, userName, password);
            //SpCsRestBreakSecurityInheritanceWeb(webUri, userName, password);
            //SpCsRestResetSecurityInheritanceWeb(webUri, userName, password);
            //SpCsRestAddUserToSecurityRoleInWeb(webUri, userName, password);
            //SpCsRestUpdateUserSecurityRoleInWeb(webUri, userName, password);
            //SpCsRestDeleteUserFromSecurityRoleInWeb(webUri, userName, password);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsRestCreateOneCommunicationSiteCollection(Uri webBaseUri,
                                                string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webBaseUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new
                    {
                        type =
                                "SP.Publishing.CommunicationSiteCreationRequest"
                    },
                    Title = "NewSiteCollectionModernCsRest",
                    Description = "NewSiteCollectionModernCsRest Description",
                    AllowFileSharingForGuestUsers = false,
                    SiteDesignId = "6142d2a0-63a5-4ba0-aede-d9fefca2c767",
                    Url = webBaseUri + "sites/NewSiteCollectionModernCsRest",
                    lcid = 1033
                };
                string endpointUrl = webBaseUri +
                                            "_api/sitepages/communicationsite/create";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 01 

        //gavdcodebegin 02
        static void SpCsRestCreateOneSiteCollection(Uri webBaseUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webBaseUri, userName, password))
            {
                string myPayload =
                    "{" +
                    "'request':  {" +
                    "'__metadata': { 'type':" +
                    "        'Microsoft.SharePoint.Portal.SPSiteCreationRequest' }," +
                    "'Title': 'NewSiteCollectionModernCsRest02'," +
                    "'Lcid': 1033," +
                    "'Description': ''," +
                    "'Classification': ''," +
                    "'ShareByEmailEnabled': false," +
                    "'SiteDesignId': '00000000-0000-0000-0000-000000000000'," +
                    "'Url': '" + webBaseUri + "/sites/NewSiteCollectionModernCsRest02'," +
                    "'WebTemplate': 'SITEPAGEPUBLISHING#0'," +
                    "'WebTemplateExtensionId': '00000000-0000-0000-0000-000000000000'" +
                    "}}";
                string endpointUrl = webBaseUri + "_api/SPSiteManager/create";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsRestCreateOneWebInSiteCollection(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.WebCreationInformation" },
                    Title = "NewWebSiteModernCsRest",
                    Description = "NewWebSiteModernCsRest Description",
                    Url = "NewWebSiteModernCsRest",
                    UseSamePermissionsAsParentSite = true,
                    WebTemplate = "STS#3"
                };
                string endpointUrl = webUri + "/_api/web/webinfos/add";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsRestReadAllSiteCollections(Uri webBaseUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webBaseUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webBaseUri +
                                "/_api/search/query?querytext='contentclass:sts_site'";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsRestReadAllWebsInSiteCollection(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/webs";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsRestUpdateOneWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.Web" },
                    Description = "NewWebSiteModernCsRest Description Updated"
                };
                string endpointUrl = subWebUri + "/_api/web";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsRestDeleteOneWebFromSiteCollection(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, headers,
                                                                    myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsRestGetRoleDefinitionsWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/roledefinitions";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsRestFindUserPermissionsWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/" +
                                            "doesuserhavepermissions(@v)?@v=" +
                                            "{'High':'2147483647', 'Low':'4294967295'}";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsRestFindOtherUserPermissionsWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/" +
                                         "getusereffectivepermissions(@v)?@v=" +
                                         "'i%3A0%23.f%7Cmembership%7C" + userName + "'";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsRestBreakSecurityInheritanceWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web" +
                                "/breakroleinheritance(copyRoleAssignments=false," +
                                "clearSubscopes=true)";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                    headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsRestResetSecurityInheritanceWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/resetroleinheritance";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsRestAddUserToSecurityRoleInWeb(Uri webUri, string userName,
                                                                    string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");

            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Find the RoleDefinitions
            int roleId = 0;
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/roledefinitions?$select=Id&" +
                                            "$filter=startswith(Name,'Full Control')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                roleId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(roleId);
            }

            // Add the User in the RoleDefinion to the List
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/lists/getbytitle" +
                          "('TestList')/items(17)/roleassignments/addroleassignment" +
                          "(principalid=" + userId + ",roledefid=" + roleId + ")";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                    headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void SpCsRestUpdateUserSecurityRoleInWeb(Uri webUri,
                                                    string userName, string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");

            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                        myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Find the RoleDefinitions
            int roleId = 0;
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/roledefinitions/getbyname" +
                                                                        "('Edit')/Id";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                        myPayload);
                roleId = int.Parse(data["d"]["Id"].ToString());
                Console.WriteLine(roleId);
            }

            // Add the User in the RoleDefinion to the List
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/" +
                            "roleassignments/addroleassignment" +
                            "(principalid=" + userId + ",roledefid=" + roleId + ")";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        static void SpCsRestDeleteUserFromSecurityRoleInWeb(Uri webUri,
                                                    string userName, string password)
        {
            Uri subWebUri = new Uri(webUri + "/NewWebSiteModernCsRest");

            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Remove the User from the List
            using (SPHttpClient client = new SPHttpClient(subWebUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = subWebUri + "/_api/web/" +
                        "roleassignments/getbyprincipalid(" +
                        "principalid=" + userId + ")";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                    headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 15
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