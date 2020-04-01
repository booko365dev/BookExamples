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

namespace LGHE
{
    class Program
    {
        static void Main(string[] args)
        {
            Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            string userName = ConfigurationManager.AppSettings["spUserName"];
            string password = ConfigurationManager.AppSettings["spUserPw"];

            //SpCsRestCreateOneListItem(webUri, userName, password);
            //SpCsRestUploadOneDocument(webUri, userName, password);
            //SpCsRestDownloadOneDocument(webUri, userName, password);
            //SpCsRestReadAllListsItems(webUri, userName, password);
            //SpCsRestReadOneListsItem(webUri, userName, password);
            //SpCsRestReadAllLibraryDocs(webUri, userName, password);
            //SpCsRestReadOneLibraryDoc(webUri, userName, password);
            //SpCsRestUpdateOneListItem(webUri, userName, password);
            //SpCsRestUpdateOneLibraryDoc(webUri, userName, password);
            //SpCsRestDeleteOneListItem(webUri, userName, password);
            //SpCsRestDeleteOneLibraryDoc(webUri, userName, password);
            //SpCsRestBreakSecurityInheritanceListItem(webUri, userName, password);
            //SpCsRestResetSecurityInheritanceListItem(webUri, userName, password);
            //SpCsRestAddUserToSecurityRoleInListItem(webUri, userName, password);
            //SpCsRestUpdateUserSecurityRoleInListItem(webUri, userName, password);
            //SpCsRestDeleteUserFromSecurityRoleInListItem(webUri, userName, password);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsRestCreateOneListItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.ListItem" },
                    Title = "NewListItemCsRest"
                };
                string endpointUrl = webUri + "/_api/web/lists/getbytitle('TestList')" +
                                                                            "/items";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 01 

        //gavdcodebegin 02
        static void SpCsRestUploadOneDocument(Uri webUri, string userName,
                                                                    string password)
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\TestDocument01.docx");
            string webUrlRel = new Uri(webUri.ToString()).AbsolutePath;

            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                Stream myPayload = System.IO.File.OpenRead(myFileInfo.FullName);
                string endpointUrl = webUri + "/_api/web/getfolderbyserverrelativeurl(" +
                                "'" + webUrlRel + "/TestLibrary')/files/add(url='" +
                                myFileInfo.Name + "',overwrite=true)";
                var data = client.ExecuteJson<Stream>(endpointUrl, HttpMethod.Post,
                                                                        myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 02 

        //gavdcodebegin 03
        static void SpCsRestDownloadOneDocument(Uri webUri, string userName,
                                                                    string password)
        {
            string webUrlRel = new Uri(webUri.ToString()).AbsolutePath;

            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/web/getfilebyserverrelativeurl(" +
                                "'" + webUrlRel + "/TestLibrary/TestDocument01.docx')" +
                                "/$value";
                Stream data = (Stream)client.ExecuteJson(endpointUrl, HttpMethod.Get,
                                                                    myPayload, true);

                byte[] result;
                using (var streamReader = new MemoryStream())
                {
                    data.CopyTo(streamReader);
                    result = streamReader.ToArray();
                }
                FileStream outputStream = new FileStream(@"C:\Temporary\TestDwload.docx",
                                    FileMode.OpenOrCreate | FileMode.Append,
                                    FileAccess.Write, FileShare.None);
                outputStream.Write(result, 0, result.Length);
                outputStream.Flush(true);
                outputStream.Close();
            }
        }
        //gavdcodeend 03 

        //gavdcodebegin 04
        static void SpCsRestReadAllListsItems(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')" +
                                                             "/items?$select=Title,Id";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsRestReadOneListsItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')" +
                                                        "/items(16)?$select=Title,Id";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsRestReadAllLibraryDocs(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestLibrary')" +
                                                             "/items?$select=Title,Id";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsRestReadOneLibraryDoc(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestLibrary')" +
                                                        "/items(22)?$select=Title,Id";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsRestUpdateOneListItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.ListItem" },
                    Title = "NewListItemCsRest_Updated"
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')/" +
                                                                            "items(16)";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsRestUpdateOneLibraryDoc(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new { type = "SP.ListItem" },
                    Title = "TestDocument01_Updated.docx"
                };
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestLibrary')/" +
                                                                        "items(22)";
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
        static void SpCsRestDeleteOneListItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')" +
                                                                    "/items(16)";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, headers,
                                                                    myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsRestDeleteOneLibraryDoc(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestLibrary')" +
                                                        "/items(22)";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, headers,
                                                                    myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsRestBreakSecurityInheritanceListItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')/" +
                    "items(17)/breakroleinheritance(copyRoleAssignments=false," +
                    "clearSubscopes=true)";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "MERGE");
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post,
                                                                    headers, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsRestResetSecurityInheritanceListItem(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri + "/_api/lists/getbytitle('TestList')/" +
                    "items(17)/resetroleinheritance";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Post, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void SpCsRestAddUserToSecurityRoleInListItem(Uri webUri, string userName,
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
        //gavdcodeend 14

        //gavdcodebegin 15
        static void SpCsRestUpdateUserSecurityRoleInListItem(Uri webUri,
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
        //gavdcodeend 15

        //gavdcodebegin 16
        static void SpCsRestDeleteUserFromSecurityRoleInListItem(Uri webUri,
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
                string endpointUrl = webUri + "/_api/web/lists/getbytitle" +
                        "('TestList')/items(17)/roleassignments/getbyprincipalid(" +
                        "principalid=" + userId + ")";
                IDictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("IF-MATCH", "*");
                headers.Add("X-HTTP-Method", "DELETE");
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
