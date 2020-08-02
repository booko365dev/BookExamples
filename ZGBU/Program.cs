using Newtonsoft.Json;
using RestSharp;
using System;
using System.Configuration;
using System.Web;

namespace ZGBU
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetTeamApp();    
            //CreateChannelApp();
            //GetChannelApp(); 
            //UpdateChannelApp();
            //DeleteChannelApp();
            //AdAppToken myTokenApp = GetAzureTokenApplication();
            //AdAppToken myTokenDel = GetAzureTokenDelegation();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 03
        static void GetTeamApp()
        {
            string graphQuery =
             "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d";

            //AdAppToken adToken = GetAzureTokenApplication();  
            AdAppToken adToken = GetAzureTokenDelegation(); 

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(graphQuery, Method.GET);
            myRequest.AddHeader("Authorization", adToken.token_type + " " + 
                                                                adToken.access_token);

            string resultText = myClient.Execute(myRequest).Content;
            Console.WriteLine(resultText);
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void CreateChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" + 
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels";

            AdAppToken adToken = GetAzureTokenApplication();
            //AdAppToken adToken = GetAzureTokenDelegation();

            string myBody = "{ " +
                                "\"displayName\": \"Graph Channel 20\"," +
                                "\"description\": \"Channel created with Graph\"" +
                            " }";

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(graphQuery, Method.POST);
            myRequest.AddHeader("Authorization", adToken.token_type + " " + 
                                                                adToken.access_token);
            myRequest.AddHeader("ContentType", "application/json");
            myRequest.AddJsonBody(myBody);

            string resultText = myClient.Execute(myRequest).Content;

            Console.WriteLine(resultText);
        }
        //gavdcodeend 04

        static void GetChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2341eebad8324e43ae86a289bff4356d@thread.tacv2";

            AdAppToken adToken = GetAzureTokenApplication();
            //AdAppToken adToken = GetAzureTokenDelegation();

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(graphQuery, Method.GET);
            myRequest.AddHeader("Authorization", adToken.token_type + " " + 
                                                                adToken.access_token);

            string resultText = myClient.Execute(myRequest).Content;

            Console.WriteLine(resultText);
        }

        //gavdcodebegin 05
        static void UpdateChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2341eebad8324e43ae86a289bff4356d@thread.tacv2";

            AdAppToken adToken = GetAzureTokenApplication();
            //AdAppToken adToken = GetAzureTokenDelegation();

            string myBody = "{ \"description\": \"Channel Description Updated\" }";

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(graphQuery, Method.PATCH);
            myRequest.AddHeader("Authorization", adToken.token_type + " " + 
                                                                adToken.access_token);
            myRequest.AddHeader("IF-MATCH", "*");
            myRequest.AddJsonBody(myBody);

            string resultText = myClient.Execute(myRequest).Content;

            Console.WriteLine(resultText);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void DeleteChannelApp()
        {
            string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                "5b409eec-a4ae-4f04-a354-0434c444265d/channels/" +
                "19:2341eebad8324e43ae86a289bff4356d@thread.tacv2";

            AdAppToken adToken = GetAzureTokenApplication();
            //AdAppToken adToken = GetAzureTokenDelegation();

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(graphQuery, Method.DELETE);
            myRequest.AddHeader("Authorization", adToken.token_type + " " + 
                                                                adToken.access_token);

            string resultText = myClient.Execute(myRequest).Content;

            Console.WriteLine(resultText);
        }
        //gavdcodeend 06

        //gavdcodebegin 01
        static AdAppToken GetAzureTokenApplication()
        {
            string LoginUrl = "https://login.microsoftonline.com";
            string ScopeUrl = "https://graph.microsoft.com/.default";

            string myClientID = ConfigurationManager.AppSettings["ClientIdApp"];
            string myClientSecret = ConfigurationManager.AppSettings["ClientSecretApp"];
            string myTenantName = ConfigurationManager.AppSettings["TenantName"];

            string myUri = LoginUrl + "/" + myTenantName + "/oauth2/v2.0/token";

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(myUri, Method.POST);
            myRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                            "grant_type=client_credentials&" +
                            "client_id=" + myClientID + "&" +
                            "client_secret=" + myClientSecret + "";
            myRequest.AddParameter("", myBody, ParameterType.RequestBody);

            string tokenJSON = myClient.Execute(myRequest).Content;
            AdAppToken tokenObj = JsonConvert.DeserializeObject<AdAppToken>(tokenJSON);

            return tokenObj;
        }
        //gavdcodeend 01

        //gavdcodebegin 07
        static AdAppToken GetAzureTokenDelegation()
        {
            string LoginUrl = "https://login.microsoftonline.com";
            string ScopeUrl = "https://graph.microsoft.com/.default";

            string myClientID = ConfigurationManager.AppSettings["ClientIdDel"];
            string myTenantName = ConfigurationManager.AppSettings["TenantName"];
            string myUserName = ConfigurationManager.AppSettings["UserName"];
            string myUserPw = ConfigurationManager.AppSettings["UserPw"];

            string myUri = LoginUrl + "/" + myTenantName + "/oauth2/v2.0/token";

            RestClient myClient = new RestClient();

            RestRequest myRequest = new RestRequest(myUri, Method.POST);
            myRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                            "grant_type=Password&" +
                            "client_id=" + myClientID + "&" +
                            "Username=" + myUserName + "&" +
                            "Password=" + myUserPw + "";
            myRequest.AddParameter("", myBody, ParameterType.RequestBody);

            string tokenJSON = myClient.Execute(myRequest).Content;
            AdAppToken tokenObj = JsonConvert.DeserializeObject<AdAppToken>(tokenJSON);

            return tokenObj;
        }
        //gavdcodeend 07
    }

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
    //gavdcodeend 02
}