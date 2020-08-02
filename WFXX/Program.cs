using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;
using System.Web;

namespace WFXX
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetTeamApp();
            //GetTeamDel();
            //CreateChannelApp();
            //CreateChannelDel();
            //GetChannelApp();   
            //GetChannelDel();   
            //UpdateChannelApp();
            //UpdateChannelDel();
            //DeleteChannelApp();
            //DeleteChannelDel();
            //AdAppToken myTokenApp = GetADTokenApplication();
            //AdAppToken myTokenDel = GetADTokenDelegation(); 

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 03
        static void GetTeamApp()
        {
            string graphQuery =
            "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                EndPoint = graphQuery,
                Method = HttpVerb.GET,
                Registration = TypeRegistration.Application
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void GetTeamDel()
        {
            string graphQuery =
            "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"],
                EndPoint = graphQuery,
                Method = HttpVerb.GET,
                Registration = TypeRegistration.Delegation
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void CreateChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "5b409eec-a4ae-4f04-a354-0434c444265d/channels";

            string myBody = "{ " +
                                "\"displayName\": \"Graph Channel 01 Application\"," +
                                "\"description\": \"Channel created with Graph\"" +
                            " }";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                EndPoint = graphQuery,
                Method = HttpVerb.POST,
                ContentType = "application/json",
                PostData = myBody,
                Registration = TypeRegistration.Application
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void CreateChannelDel()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "5b409eec-a4ae-4f04-a354-0434c444265d/channels";

            string myBody = "{ " +
                                "\"displayName\": \"Graph Channel 02 Delegation\"," +
                                "\"description\": \"Channel created with Graph\"" +
                            " }";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"],
                EndPoint = graphQuery,
                Method = HttpVerb.POST,
                ContentType = "application/json",
                PostData = myBody,
                Registration = TypeRegistration.Delegation
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 06

        static void GetChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:621a61495881425791c65ba19ccb411b@thread.tacv2";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                EndPoint = graphQuery,
                Method = HttpVerb.GET,
                Registration = TypeRegistration.Application
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }

        static void GetChannelDel()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2d2d0b00fbb74411873da4a9dd2f0d7c@thread.tacv2";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"],
                EndPoint = graphQuery,
                Method = HttpVerb.GET,
                Registration = TypeRegistration.Delegation
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }

        //gavdcodebegin 07
        static void UpdateChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:621a61495881425791c65ba19ccb411b@thread.tacv2";

            string myBody = "{ \"description\": \"Channel Description Updated\" }";

            List<HeaderConfig> myHeadersList = new List<HeaderConfig>();
            HeaderConfig myHeaderMat = new HeaderConfig
            {
                HeaderTitle = "IF-MATCH",
                HeaderValue = "*"
            };
            myHeadersList.Add(myHeaderMat);

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                EndPoint = graphQuery,
                Method = HttpVerb.PATCH,
                ContentType = "application/json",
                Headers = myHeadersList,
                PostData = myBody,
                Registration = TypeRegistration.Application
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void UpdateChannelDel()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2d2d0b00fbb74411873da4a9dd2f0d7c@thread.tacv2";

            string myBody = "{ \"description\": \"Channel Description Updated\" }";

            List<HeaderConfig> myHeadersList = new List<HeaderConfig>();
            HeaderConfig myHeaderMat = new HeaderConfig
            {
                HeaderTitle = "IF-MATCH",
                HeaderValue = "*"
            };
            myHeadersList.Add(myHeaderMat);

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"],
                EndPoint = graphQuery,
                Method = HttpVerb.PATCH,
                ContentType = "application/json",
                Headers = myHeadersList,
                PostData = myBody,
                Registration = TypeRegistration.Delegation
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void DeleteChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:621a61495881425791c65ba19ccb411b@thread.tacv2";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                EndPoint = graphQuery,
                Method = HttpVerb.DELETE,
                Registration = TypeRegistration.Application
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void DeleteChannelDel()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2d2d0b00fbb74411873da4a9dd2f0d7c@thread.tacv2";

            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"],
                EndPoint = graphQuery,
                Method = HttpVerb.DELETE,
                Registration = TypeRegistration.Delegation
            };

            Tuple<string, string> resultText = myClient.SendGraphRequest();

            Console.WriteLine(resultText.Item1);
            Console.WriteLine(resultText.Item2);
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static AdAppToken GetADTokenApplication()
        {
            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdApp"],
                ClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"],
                TenantName = ConfigurationManager.AppSettings["TenantName"]
            };

            AdAppToken resultToken = myClient.GetAzureTokenApplication();

            return resultToken;
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static AdAppToken GetADTokenDelegation()
        {
            RestGraphClient myClient = new RestGraphClient
            {
                ClientID = ConfigurationManager.AppSettings["ClientIdDel"],
                TenantName = ConfigurationManager.AppSettings["TenantName"],
                UserName = ConfigurationManager.AppSettings["UserName"],
                UserPw = ConfigurationManager.AppSettings["UserPw"]
            };

            AdAppToken resultToken = myClient.GetAzureTokenDelegation();

            return resultToken;
        }
        //gavdcodeend 12
    }

    //gavdcodebegin 01
    public class RestGraphClient
    {
        public string ClientID { get; set; }
        public string ClientSecret { get; set; }
        public string TenantName { get; set; }
        public string EndPoint { get; set; }
        public HttpVerb Method { get; set; }
        public TypeRegistration Registration { get; set; }
        public string ContentType { get; set; }
        public string PostData { get; set; }
        public string UserName { get; set; }
        public string UserPw { get; set; }
        public List<HeaderConfig> Headers { get; set; }

        public RestGraphClient()
        {
        }

        public Tuple<string, string> SendGraphRequest()
        {
            AdAppToken adToken = new AdAppToken();
            if (Registration == TypeRegistration.Application)
                adToken = GetAzureTokenApplication();
            else if (Registration == TypeRegistration.Delegation)
                adToken = GetAzureTokenDelegation();

            if (adToken != null)
            {
                List<HeaderConfig> myHeadersList = new List<HeaderConfig>();
                HeaderConfig authorizationHeader = new HeaderConfig
                {
                    HeaderTitle = "Authorization",
                    HeaderValue = adToken.token_type + " " + adToken.access_token
                };
                myHeadersList.Add(authorizationHeader);
                Headers = myHeadersList;

                return SendGraphRequestInternal();
            }
            else
            {
                Tuple<string, string> tplReturn = new Tuple<string, string>
                                    ("Error", string.Empty);
                return tplReturn;
            }
        }

        private Tuple<string, string> SendGraphRequestInternal()
        {
            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(EndPoint);

            myRequest.Method = Method.ToString();
            myRequest.ContentLength = 0;
            myRequest.ContentType = ContentType;
            if (Headers != null)
            {
                foreach (HeaderConfig oneHeader in Headers)
                {
                    myRequest.Headers.Add(oneHeader.HeaderTitle, oneHeader.HeaderValue);
                }
            }

            if (string.IsNullOrEmpty(PostData) == false)
            {
                byte[] bodyBytes = Encoding.GetEncoding("iso-8859-1").GetBytes(PostData);
                myRequest.ContentLength = bodyBytes.Length;

                using (Stream writeStream = myRequest.GetRequestStream())
                {
                    writeStream.Write(bodyBytes, 0, bodyBytes.Length);
                }
            }

            try
            {
                using (HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse())
                {
                    string responseValue = string.Empty;

                    using (Stream responseStream = myResponse.GetResponseStream())
                    {
                        if (responseStream != null)
                            using (StreamReader myReader = new StreamReader(responseStream))
                            {
                                responseValue = myReader.ReadToEnd();
                            }
                    }

                    Tuple<string, string> tplReturn = new Tuple<string, string>
                                        (myResponse.StatusCode.ToString(), responseValue);
                    return tplReturn;
                }
            }
            catch (Exception ex)
            {
                Tuple<string, string> tplReturn = new Tuple<string, string>
                                    ("Error", ex.ToString());
                return tplReturn;
            }
        }

        public AdAppToken GetAzureTokenApplication()
        {
            string LoginUrl = "https://login.microsoftonline.com";
            string ScopeUrl = "https://graph.microsoft.com/.default";

            string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";
            string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                "grant_type=client_credentials&" +
                "client_id=" + ClientID + "&" +
                "client_secret=" + ClientSecret + "";

            RestGraphClient myClient = new RestGraphClient
            {
                EndPoint = myUri,
                Method = HttpVerb.POST,
                ContentType = "application/x-www-form-urlencoded",
                PostData = myBody
            };

            Tuple<string, string> tokenJSON = myClient.SendGraphRequestInternal();
            if (tokenJSON.Item1.Contains("Error") == false)
            {
                AdAppToken tokenObj =
                            JsonConvert.DeserializeObject<AdAppToken>(tokenJSON.Item2);
                return tokenObj;
            }

            return null;
        }

        public AdAppToken GetAzureTokenDelegation()
        {
            string LoginUrl = "https://login.microsoftonline.com";
            string ScopeUrl = "https://graph.microsoft.com/.default";

            string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";
            string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                            "grant_type=Password&" +
                            "client_id=" + ClientID + "&" +
                            "Username=" + UserName + "&" +
                            "Password=" + UserPw + "";

            RestGraphClient myClient = new RestGraphClient
            {
                EndPoint = myUri,
                Method = HttpVerb.POST,
                ContentType = "application/x-www-form-urlencoded",
                PostData = myBody
            };

            Tuple<string, string> tokenJSON = myClient.SendGraphRequestInternal();
            if (tokenJSON.Item1.Contains("Error") == false)
            {
                AdAppToken tokenObj =
                            JsonConvert.DeserializeObject<AdAppToken>(tokenJSON.Item2);
                return tokenObj;
            }

            return null;
        }
    }
    //gavdcodeend 01

    //gavdcodebegin 02
    public class AdAppToken
    {
        public string token_type { get; set; }
        public string expires_in { get; set; }
        public string ext_expires_in { get; set; }
        public string expires_on { get; set; }
        public string not_before { get; set; }
        public string resource { get; set; }
        public string access_token { get; set; }
    }

    public enum HttpVerb
    {
        GET,
        PATCH,
        POST,
        PUT,
        DELETE
    }

    public enum TypeRegistration
    {
        Application,
        Delegation
    }

    public class HeaderConfig
    {
        public string HeaderTitle { get; set; }
        public string HeaderValue { get; set; }
    }
    //gavdcodeend 02
}
