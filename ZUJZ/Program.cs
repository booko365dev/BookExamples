using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Security;

namespace ZUJZ
{
    class Program
    {
        static void Main(string[] args)
        {
            //SpCsRestCreateOneList();
            //SpCsRestReadeAllLists();
            //SpCsRestReadeOneList();
            //SpCsRestUpdateOneList();
            //SpCsRestDeleteOneList();
            //SpCsRestAddOneFieldToList();
            //SpCsRestReadAllFieldsFromList();
            //SpCsRestReadOneFieldFromList();
            //SpCsRestUpdateOneFieldInList();
            //SpCsRestDeleteOneFieldFromList();
            //SpCsRestBreakSecurityInheritanceList();
            //SpCsRestResetSecurityInheritanceList();
            //SpCsRestAddUserToSecurityRoleInList();
            //SpCsRestUpdateUserSecurityRoleInList();
            //SpCsRestDeleteUserFromSecurityRoleInList();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void SpCsRestCreateOneList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/web/lists";
            string PostRestQuery = "{ '__metadata': { 'type': 'SP.List' }, " +
                                    "'Title': 'NewListRest', " +
                                    "'BaseTemplate': 100, " +
                                    "'Description': 'New List created using REST' }";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.POST);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestReadeAllLists()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists?$select=Title,Id";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestReadeOneList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestUpdateOneList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')";
            string PostRestQuery = "{ '__metadata': { 'type': 'SP.List' }, " +
                                    "'Description': 'New List Description' }";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.MERGE);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestDeleteOneList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.DELETE);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestAddOneFieldToList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/fields";
            string PostRestQuery = "{ '__metadata': { 'type': 'SP.Field' }, " +
                                    "'Title': 'MyMultilineField', " +
                                    "'FieldTypeKind': 3 }";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.POST);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestReadAllFieldsFromList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/fields";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestReadOneFieldFromList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/fields/" +
                                        "getbytitle('MyMultilineField')";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestUpdateOneFieldInList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/fields/" +
                "                               getbytitle('MyMultilineField')";
            string PostRestQuery = "{ '__metadata': { 'type': 'SP.Field' }, " +
                                    "'Description': 'New Field Description' }";
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.MERGE);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestDeleteOneFieldFromList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/fields/" +
                "                           getbytitle('MyMultilineField')";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.DELETE);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestBreakSecurityInheritanceList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/" +
                "breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.POST);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestResetSecurityInheritanceList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string BaseRestQuery = "/_api/lists/getbytitle('NewListRest')/" +
                "resetroleinheritance";
            string PostRestQuery = string.Empty;
            string ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.POST);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestAddUserToSecurityRoleInList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string PostRestQuery = string.Empty;
            string BaseRestQuery = string.Empty;
            string ResponseResult = string.Empty;
            JObject resultJson = null;

            // Find the User
            BaseRestQuery = "/_api/web/siteusers?$select=Id&" +
                                        "$filter=startswith(Title,'MOD')";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);
            resultJson = JObject.Parse(ResponseResult);
            int userId = int.Parse(resultJson["results"][0]["Id"].ToString());

            // Find the RoleDefinitions
            BaseRestQuery = "/_api/web/roledefinitions?$select=Id&" +
                                        "$filter=startswith(Name,'Full Control')";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);
            resultJson = JObject.Parse(ResponseResult);
            int roleId = int.Parse(resultJson["results"][0]["Id"].ToString());

            // Add the User in the RoleDefinion to the List
            BaseRestQuery = "/_api/web/lists/getbytitle('NewListRest')/roleassignments/" +
                "addroleassignment(principalid=" + userId + ",roledefid=" + roleId + ")";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.POST);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestUpdateUserSecurityRoleInList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string PostRestQuery = string.Empty;
            string BaseRestQuery = string.Empty;
            string ResponseResult = string.Empty;
            JObject resultJson = null;

            // Find the User
            BaseRestQuery = "/_api/web/siteusers/?$select=Id&" +
                                        "$filter=startswith(Title,'MOD')";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);
            resultJson = JObject.Parse(ResponseResult);
            int userId = int.Parse(resultJson["results"][0]["Id"].ToString());

            // Find the RoleDefinitions
            BaseRestQuery = "/_api/web/roledefinitions/getbyname('Edit')/id";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);
            resultJson = JObject.Parse(ResponseResult);
            int roleId = int.Parse(resultJson["Id"].ToString());

            // Add the User in the RoleDefinion to the List
            BaseRestQuery = "/_api/web/lists/getbytitle('NewListRest')/roleassignments/" +
                "addroleassignment(principalid=" + userId + ",roledefid=" + roleId + ")";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.MERGE);

            Console.WriteLine(ResponseResult);
        }

        static void SpCsRestDeleteUserFromSecurityRoleInList()
        {
            string SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            string PostRestQuery = string.Empty;
            string BaseRestQuery = string.Empty;
            string ResponseResult = string.Empty;
            JObject resultJson = null;

            // Find the User
            BaseRestQuery = "/_api/web/siteusers/?$select=Id&" +
                                        "$filter=startswith(Title,'MOD')";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.GET);
            resultJson = JObject.Parse(ResponseResult);
            int userId = int.Parse(resultJson["results"][0]["Id"].ToString());

            // Remove the User from the List
            BaseRestQuery = "/_api/web/lists/GetByTitle('NewListRest')/roleassignments/" +
                "getbyprincipalid(principalid=" + userId + ")";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                                     BaseRestQuery,
                                                     PostRestQuery,
                                                     TypeRequest.DELETE);

            Console.WriteLine(ResponseResult);
        }

        //----------------------------------------------------------------------------------------
        /// <summary>
        /// Executes a GET call to the REST method
        /// </summary>
        /// <param name="SiteBaseUrl">https://dominio.sharepoint.com/sites/namesite</param>
        /// <param name="BaseRestQuery">/_api/web/created</param>
        /// <returns></returns>
        public static string ExecuteRestQuery(string SiteBaseUrl,
                                                 string BaseRestQuery)
        {
            string SiteRestUrl = SiteBaseUrl + BaseRestQuery;

            SharePointOnlineCredentials myCredentials = GetCredentials();

            using (WebClient myWebClient = new WebClient())
            {
                myWebClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                myWebClient.Credentials = myCredentials;
                myWebClient.Headers.Add(HttpRequestHeader.ContentType,
                                        "application/json;odata=verbose");
                myWebClient.Headers.Add(HttpRequestHeader.Accept,
                                        "application/json;odata=verbose");

                string myResult = myWebClient.DownloadString(SiteRestUrl);

                JObject resultJson = JObject.Parse(myResult);
                return resultJson["d"].ToString();
            }
        }

        /// <summary>
        /// Executes a GET, POST, MERGE or DELETE call to the REST method
        /// </summary>
        /// <param name="SiteBaseUrl">https://dominio.sharepoint.com/sites/namesite</param>
        /// <param name="BaseRestQuery">/_api/web/lists</param>
        /// <param name="PostRestQuery">{ '__metadata': { 'type': 'SP.List' }, ...</param>
        /// <param name="RequestType">TypeRequest.POST / GET / MERGE / DELETE</param>
        /// <returns></returns>
        public static string ExecuteRestQuery(string SiteBaseUrl,
                                              string BaseRestQuery,
                                              string PostRestQuery,
                                              TypeRequest RequestType)
        {
            if (RequestType.Equals(TypeRequest.GET) == true)
            { return ExecuteRestQuery(SiteBaseUrl, BaseRestQuery); }

            string SiteRestUrl = SiteBaseUrl + BaseRestQuery;

            CookieContainer myCookies = GetAuthCookies(new Uri(SiteBaseUrl));
            string myFormDigest = GetFormDigest(SiteBaseUrl, myCookies);

            HttpWebRequest myWebReq = GetRequest(SiteRestUrl, myCookies,
                                                PostRestQuery.Length);
            myWebReq.Headers.Add("X-RequestDigest", myFormDigest);
            switch (RequestType)
            {
                case TypeRequest.MERGE:
                    myWebReq.Headers.Add("IF-MATCH", "*");
                    myWebReq.Headers.Add("X-HTTP-Method", "MERGE");
                    break;
                case TypeRequest.DELETE:
                    myWebReq.Headers.Add("IF-MATCH", "*");
                    myWebReq.Headers.Add("X-HTTP-Method", "DELETE");
                    break;
            }

            StreamWriter myReqStream = new StreamWriter(myWebReq.GetRequestStream());
            myReqStream.Write(PostRestQuery);
            myReqStream.Flush();

            JObject resultJson = GetResult(myWebReq);
            if (resultJson != null)
            { return resultJson["d"].ToString(); }
            else
            { return string.Empty; }
        }

        private static CookieContainer GetAuthCookies(Uri SiteBaseUrl)
        {
            SharePointOnlineCredentials myCredentials = GetCredentials();

            string authCookie = myCredentials.GetAuthenticationCookie(SiteBaseUrl);
            CookieContainer myCookies = new CookieContainer();
            myCookies.SetCookies(SiteBaseUrl, authCookie);

            return myCookies;
        }

        private static SharePointOnlineCredentials GetCredentials()
        {
            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            SharePointOnlineCredentials myCredentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return myCredentials;
        }

        private static string GetFormDigest(string SiteBaseUrl, CookieContainer Cookies)
        {
            string resourceUrl = SiteBaseUrl + "/_api/contextinfo";
            HttpWebRequest myWebReq = GetRequest(resourceUrl, Cookies, 0);

            JObject resultJson = GetResult(myWebReq);
            return resultJson["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
        }

        private static HttpWebRequest GetRequest(string ReqUrl, CookieContainer Cookies,
                        long ContentLenght)
        {
            HttpWebRequest myWebReq = (HttpWebRequest)HttpWebRequest.Create(ReqUrl);
            myWebReq.CookieContainer = Cookies;
            myWebReq.Method = "POST";
            myWebReq.Accept = "application/json;odata=verbose";
            myWebReq.ContentLength = ContentLenght;
            myWebReq.ContentType = "application/json;odata=verbose";

            return myWebReq;
        }

        private static JObject GetResult(HttpWebRequest WebRequest)
        {
            string myResult = string.Empty;
            WebResponse myWebResp = WebRequest.GetResponse();
            using (StreamReader myRespStream = new StreamReader
                                                (myWebResp.GetResponseStream()))
            { myResult = myRespStream.ReadToEnd(); }

            if (string.IsNullOrEmpty(myResult) == false)
            {
                JObject resultJson = JObject.Parse(myResult);
                return resultJson;
            }
            else
            {
                JObject resultJson = null;
                return resultJson;
            }
        }

        public enum TypeRequest { GET, POST, MERGE, DELETE }
    }
}

