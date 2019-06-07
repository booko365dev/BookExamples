using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Security;

namespace EJON
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginCsom();
            Web rootWeb = spCtx.Web;
            spCtx.Load(rootWeb);
            spCtx.ExecuteQuery();
            Console.WriteLine(rootWeb.Created.ToShortDateString());

            ClientContext spPnpCtx = LoginPnPCore();
            Web rootWebPnp = spPnpCtx.Web;
            spPnpCtx.Load(rootWebPnp);
            spPnpCtx.ExecuteQuery();
            Console.WriteLine(rootWebPnp.Created.ToShortDateString());

            LoginPnPCoreDirectly();

            string SiteBaseUrl    = string.Empty;
            string BaseRestQuery  = string.Empty;
            string PostRestQuery  = string.Empty;
            string ResponseResult = string.Empty;

            SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            BaseRestQuery = "/_api/web/created";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl, 
                                              BaseRestQuery);
            Console.WriteLine(ResponseResult);

            SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];
            BaseRestQuery = "/_api/web/created";
            ResponseResult = ExecuteRestQuery(SiteBaseUrl,
                                              BaseRestQuery,
                                              null,
                                              TypeRequest.GET);
            Console.WriteLine(ResponseResult);

            //SiteBaseUrl = ConfigurationManager.AppSettings["spUrl"];  //==> To create a new List
            //BaseRestQuery = "/_api/web/lists";
            //PostRestQuery = "{ '__metadata': { 'type': 'SP.List' }, 'AllowContentTypes': true, 'BaseTemplate': 100, 'ContentTypesEnabled': true, 'Description': 'My gavd list description', 'Title': 'NewListRest' }";
            //ResponseResult = ExecuteRestQuery(SiteBaseUrl, 
            //                                  BaseRestQuery, 
            //                                  PostRestQuery,
            //                                  TypeRequest.POST);
            //Console.WriteLine(ResponseResult);

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

