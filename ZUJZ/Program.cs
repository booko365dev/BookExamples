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

namespace ZUJZ
{
    class Program
    {
        static void Main(string[] args)
        {
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            //SpCsRest_CreateOneList(webUri, userName, password);
            //SpCsRest_ReadeAllLists(webUri, userName, password);
            //SpCsRest_ReadeOneList(webUri, userName, password);
            //SpCsRest_UpdateOneList(webUri, userName, password);
            //SpCsRest_DeleteOneList(webUri, userName, password);
            //SpCsRest_AddOneFieldToList(webUri, userName, password);
            //SpCsRest_ReadAllFieldsFromList(webUri, userName, password);
            //SpCsRest_ReadOneFieldFromList(webUri, userName, password);
            //SpCsRest_UpdateOneFieldInList(webUri, userName, password);
            //SpCsRest_DeleteOneFieldFromList(webUri, userName, password);
            //SpCsRest_BreakSecurityInheritanceList(webUri, userName, password);
            //SpCsRest_ResetSecurityInheritanceList(webUri, userName, password);
            //SpCsRest_AddUserToSecurityRoleInList(webUri, userName, password);
            //SpCsRest_UpdateUserSecurityRoleInList(webUri, userName, password);
            //SpCsRest_DeleteUserFromSecurityRoleInList(webUri, userName, password);
            //SpCsRest_ColumnIndex(webUri, userName, password);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsRest_CreateOneList(Uri webUri, string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.List" },
                    Title = "NewListRestCs",
                    BaseTemplate = 100,
                    Description = "Test NewListRestCs",
                    AllowContentTypes = true,
                    ContentTypesEnabled = true
                };
                string endpointUrl = webUri + "/_api/web/lists";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 01 

        //gavdcodebegin 02
        static void SpCsRest_ReadeAllLists(Uri webUri, string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists?$select=Title,Id";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsRest_ReadeOneList(Uri webUri, string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsRest_UpdateOneList(Uri webUri, string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.List" },
                    Description = "New List Description"
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsRest_DeleteOneList(Uri webUri, string userName, string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsRest_AddOneFieldToList(Uri webUri, string userName,
                                                                string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.Field" },
                    Title = "MyMultilineField",
                    FieldTypeKind = 3
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                                                            "/fields";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsRest_ReadAllFieldsFromList(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                                                            "/fields";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsRest_ReadOneFieldFromList(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                              "/fields/getbytitle('MyMultilineField')";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsRest_UpdateOneFieldInList(Uri webUri, string userName,
                                                                string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.Field" },
                    Description = "New Field Description"
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                                "/fields/getbytitle('MyMultilineField')";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsRest_DeleteOneFieldFromList(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                                "/fields/getbytitle('MyMultilineField')";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsRest_BreakSecurityInheritanceList(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')/" +
                    "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)";
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
        static void SpCsRest_ResetSecurityInheritanceList(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')/" +
                    "resetroleinheritance";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsRest_AddUserToSecurityRoleInList(Uri webUri, string userName,
                                                                    string password)
        {
            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Find the RoleDefinitions
            int roleId = 0;
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/roledefinitions?$select=Id&" +
                                            "$filter=startswith(Name,'Full Control')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                roleId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(roleId);
            }

            // Add the User in the RoleDefinion to the List
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/lists/getbytitle" +
                          "('NewListRestCs')/roleassignments/addroleassignment" +
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
        static void SpCsRest_UpdateUserSecurityRoleInList(Uri webUri,
                                                    string userName, string password)
        {
            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                        myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Find the RoleDefinitions
            int roleId = 0;
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/roledefinitions/getbyname" +
                                                                        "('Edit')/Id";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                        myPayload);
                roleId = int.Parse(data["d"]["Id"].ToString());
                Console.WriteLine(roleId);
            }

            // Add the User in the RoleDefinion to the List
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/lists/getbytitle" +
                            "('NewListRestCs')/roleassignments/addroleassignment" +
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
        static void SpCsRest_DeleteUserFromSecurityRoleInList(Uri webUri,
                                                    string userName, string password)
        {
            // Find the User
            int userId = 0;
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'MOD')";
                var data = (JObject)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                            myPayload);
                userId = int.Parse(data["d"]["results"][0]["Id"].ToString());
                Console.WriteLine(userId);
            }

            // Remove the User from the List
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/lists/GetByTitle" +
                        "('NewListRestCs')/roleassignments/getbyprincipalid(principalid=" +
                        userId + ")";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                    headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 15

        //gavdcodebegin 16
        static void SpCsRest_ColumnIndex(Uri webUri, string userName,
                                                                string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.Field" },
                    Indexed = "true"
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('NewListRestCs')" +
                                                "/fields/getbytitle('MyMultilineField')";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 16
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
