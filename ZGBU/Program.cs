using Newtonsoft.Json;
using RestSharp;
using System.Configuration;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static AdAppToken GetAzureTokenApplication(string TenantName, string ClientId, 
                                                                    string ClientSecret)
{
    string LoginUrl = "https://login.microsoftonline.com";
    string ScopeUrl = "https://graph.microsoft.com/.default";

    string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(myUri, Method.Post);
    myRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");

    string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                    "grant_type=client_credentials&" +
                    "client_id=" + ClientId + "&" +
                    "client_secret=" + ClientSecret + "";
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string tokenJSON = myClient.ExecuteAsync(myRequest).Result.Content;
    AdAppToken tokenObj = JsonConvert.DeserializeObject<AdAppToken>(tokenJSON);

    return tokenObj;
}
//gavdcodeend 01

//gavdcodebegin 07
static AdAppToken GetAzureTokenDelegation(string TenantName, string ClientId,
                                                         string UserName, string UserPw)
{
    string LoginUrl = "https://login.microsoftonline.com";
    string ScopeUrl = "https://graph.microsoft.com/.default";

    string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(myUri, Method.Post);
    myRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");

    string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                    "grant_type=Password&" +
                    "client_id=" + ClientId + "&" +
                    "Username=" + UserName + "&" +
                    "Password=" + UserPw + "";
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string tokenJSON = myClient.ExecuteAsync(myRequest).Result.Content;
    AdAppToken tokenObj = JsonConvert.DeserializeObject<AdAppToken>(tokenJSON);

    return tokenObj;
}
//gavdcodeend 07

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 03
static void GetTeamApp()
{
    string graphQuery =
     "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c";

    AdAppToken adToken = GetAzureTokenApplication(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;
    Console.WriteLine(resultText);
}
//gavdcodeend 03

//gavdcodebegin 04
static void CreateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels";

    AdAppToken adToken = GetAzureTokenApplication(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel 20\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Post);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("ContentType", "application/json");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 04

static void GetChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenApplication(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}

//gavdcodebegin 05
static void UpdateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenApplication(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Patch);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("IF-MATCH", "*");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 05

//gavdcodebegin 06
static void DeleteChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenApplication(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Delete);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 06

//gavdcodebegin 08
static void GetTeamDel()
{
    string graphQuery =
     "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c";

    AdAppToken adToken = GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;
    Console.WriteLine(resultText);
}
//gavdcodeend 08

//gavdcodebegin 09
static void CreateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels";

    AdAppToken adToken = GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel 20\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Post);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("ContentType", "application/json");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 09

static void GetChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}

//gavdcodebegin 10
static void UpdateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Patch);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("IF-MATCH", "*");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 10

//gavdcodebegin 11
static void DeleteChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new RestClient();

    RestRequest myRequest = new RestRequest(graphQuery, Method.Delete);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 11

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//GetTeamApp();    
//CreateChannelApp();
//GetChannelApp(); 
//UpdateChannelApp();
//DeleteChannelApp();
//AdAppToken myToken = GetAzureTokenApplication(
//    ConfigurationManager.AppSettings["TenantName"],
//    ConfigurationManager.AppSettings["ClientIdWithSecret"],
//    ConfigurationManager.AppSettings["ClientSecret"]); Console.WriteLine(myToken.access_token);
//GetTeamDel();    
//CreateChannelDel();
//GetChannelDel(); 
//UpdateChannelDel();
//DeleteChannelDel();
//AdAppToken myToken = GetAzureTokenDelegation(
//    ConfigurationManager.AppSettings["TenantName"],
//    ConfigurationManager.AppSettings["ClientIdWithAccPw"],
//    ConfigurationManager.AppSettings["UserName"],
//    ConfigurationManager.AppSettings["UserPw"]); Console.WriteLine(myToken.access_token);

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

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

#nullable enable