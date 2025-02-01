using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using RestSharp;
using System.Configuration;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Cryptography.X509Certificates;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static AdAppToken CsRestSharp_GetAzureTokenApplicationSecret(string TenantName,
                                                  string ClientId, string ClientSecret)
{
    string LoginUrl = "https://login.microsoftonline.com";
    string ScopeUrl = "https://graph.microsoft.com/.default";

    string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";

    RestClient myClient = new();

    RestRequest myRequest = new(myUri, Method.Post);
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
//gavdcodeend 001

//gavdcodebegin 012
static AdAppToken CsRestSharp_GetAzureTokenApplicationCertificate(string TenantName,
                string ClientId, string CertificateFilePath, string CertificateFilePw,
                string CertificateThumbprint)
{
    string LoginUrl = "https://login.microsoftonline.com";
    string ScopeUrl = "https://graph.microsoft.com/.default";

    string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";


    X509Certificate2 myCertificate = new(CertificateFilePath, CertificateFilePw);
    RestClientOptions clientOptions = new(myUri)
    {
        ClientCertificates = new X509CertificateCollection() { myCertificate }
    };
    RestClient myClient = new(clientOptions);

    RestRequest myRequest = new(myUri, Method.Post);
    myRequest.AddHeader("Content-Type", "application/x-www-form-urlencoded");

    // Get the Assertion using the Certificate File
    //string clientAssertion = GenerateClientAssertionWithFile(TenantName, ClientId, 
    //                                            CertificateFilePath, CertificateFilePw);

    // Get the Assertion using the Certificate Thumbprint
    string clientAssertion = GenerateClientAssertionWithThumbprint(TenantName, ClientId,
                                                CertificateThumbprint);

    string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
        "grant_type=client_credentials&" +
        "client_assertion_type=urn:ietf:params:oauth:client-assertion-type:jwt-bearer&" +
        "client_assertion=" + clientAssertion;

    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string tokenJSON = myClient.ExecuteAsync(myRequest).Result.Content;
    AdAppToken tokenObj = JsonConvert.DeserializeObject<AdAppToken>(tokenJSON);

    return tokenObj;
}
//gavdcodeend 012

//gavdcodebegin 013
static string GenerateClientAssertionWithFile(string TenantName, string ClientId, 
                                string CertificateFilePath, string CertificateFilePw)
{
    X509Certificate2 myCertificate = new(CertificateFilePath, CertificateFilePw);

    // Create the JWT header
    JwtHeader myHeader = new(
        new SigningCredentials(
            new X509SecurityKey(myCertificate), SecurityAlgorithms.RsaSha256)) { 
                {"x5t", Convert.ToBase64String(myCertificate.GetCertHash())} 
    };

    // Create the JWT payload
    JwtPayload myPayload = new()
    { 
        {"aud", $"https://login.microsoftonline.com/{TenantName}/v2.0"}, 
        {"iss", ClientId}, 
        {"sub", ClientId}, 
        {"jti", Guid.NewGuid().ToString()}, 
        {"exp", new DateTimeOffset(DateTime.UtcNow.AddMinutes(10)).ToUnixTimeSeconds()}, 
        {"nbf", new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds()} 
    };

    // Create the JWT token
    JwtSecurityToken jwtToken = new(myHeader, myPayload);

    // Encode the JWT token to create the client_assertion
    JwtSecurityTokenHandler tokenHandler = new(); 
    
    return tokenHandler.WriteToken(jwtToken); 
}

static string GenerateClientAssertionWithThumbprint(string TenantName, string ClientId,
                                string CertificateThumbprint)
{
    X509Store myStore = new(StoreName.My, StoreLocation.CurrentUser);
    myStore.Open(OpenFlags.ReadOnly);
    X509Certificate2 myCertificate = myStore.Certificates
        .Find(X509FindType.FindByThumbprint, CertificateThumbprint, false)
        .OfType<X509Certificate2>()
        .FirstOrDefault();
    myStore.Close();

    if (myCertificate == null)
    {
        throw new Exception("Certificate not found");
    }

    // Create the JWT header
    JwtHeader myHeader = new(
        new SigningCredentials(
            new X509SecurityKey(myCertificate), SecurityAlgorithms.RsaSha256)) {
                {"x5t", Convert.ToBase64String(myCertificate.GetCertHash())}
    };

    // Create the JWT payload
    JwtPayload myPayload = new()
    {
        {"aud", $"https://login.microsoftonline.com/{TenantName}/v2.0"},
        {"iss", ClientId},
        {"sub", ClientId},
        {"jti", Guid.NewGuid().ToString()},
        {"exp", new DateTimeOffset(DateTime.UtcNow.AddMinutes(10)).ToUnixTimeSeconds()},
        {"nbf", new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds()}
    };

    // Create the JWT token
    JwtSecurityToken jwtToken = new(myHeader, myPayload);

    // Encode the JWT token to create the client_assertion
    JwtSecurityTokenHandler tokenHandler = new();

    return tokenHandler.WriteToken(jwtToken);
}
//gavdcodeend 013

//gavdcodebegin 007
static AdAppToken CsRestSharp_GetAzureTokenDelegation(string TenantName,
                                        string ClientId, string UserName, string UserPw)
{
    string LoginUrl = "https://login.microsoftonline.com";
    string ScopeUrl = "https://graph.microsoft.com/.default";

    string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";

    RestClient myClient = new();

    RestRequest myRequest = new(myUri, Method.Post);
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
//gavdcodeend 007

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 003
static void CsRestSharp_GetTeamApp()
{
    string graphQuery =
     "https://graph.microsoft.com/v1.0/teams/dd1223a2-28a7-47d4-afc2-f42eae94f037";

    AdAppToken adToken = CsRestSharp_GetAzureTokenApplicationSecret(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;
    Console.WriteLine(resultText);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsRestSharp_CreateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "dd1223a2-28a7-47d4-afc2-f42eae94f037";

    AdAppToken adToken = CsRestSharp_GetAzureTokenApplicationSecret(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Post);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("ContentType", "application/json");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.Execute(myRequest).Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 004

static void CsRestSharp_GetChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenApplicationSecret(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}

//gavdcodebegin 005
static void CsRestSharp_UpdateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenApplicationSecret(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Patch);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("IF-MATCH", "*");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsRestSharp_DeleteChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenApplicationSecret(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithSecret"],
                                ConfigurationManager.AppSettings["ClientSecret"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Delete);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 006

//gavdcodebegin 008
static void CsRestSharp_GetTeamDel()
{
    string graphQuery =
     "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c";

    AdAppToken adToken = CsRestSharp_GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;
    Console.WriteLine(resultText);
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsRestSharp_CreateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels";

    AdAppToken adToken = CsRestSharp_GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel 20\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Post);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("ContentType", "application/json");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 009

static void CsRestSharp_GetChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Get);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}

//gavdcodebegin 010
static void CsRestSharp_UpdateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Patch);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);
    myRequest.AddHeader("IF-MATCH", "*");
    myRequest.AddParameter("", myBody, ParameterType.RequestBody);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsRestSharp_DeleteChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:a82727613c6d4a679233d153c40c7fb7@thread.tacv2";

    AdAppToken adToken = CsRestSharp_GetAzureTokenDelegation(
                                ConfigurationManager.AppSettings["TenantName"],
                                ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

    RestClient myClient = new();

    RestRequest myRequest = new(graphQuery, Method.Delete);
    myRequest.AddHeader("Authorization", adToken.token_type + " " +
                                                        adToken.access_token);

    string resultText = myClient.ExecuteAsync(myRequest).Result.Content;

    Console.WriteLine(resultText);
}
//gavdcodeend 011

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 013 ***

//CsRestSharp_GetTeamApp();    
//CsRestSharp_CreateChannelApp();
//CsRestSharp_GetChannelApp(); 
//CsRestSharp_UpdateChannelApp();
//CsRestSharp_DeleteChannelApp();
//AdAppToken myToken = CsRestSharp_GetAzureTokenApplicationSecret(
//    ConfigurationManager.AppSettings["TenantName"],
//    ConfigurationManager.AppSettings["ClientIdWithSecret"],
//    ConfigurationManager.AppSettings["ClientSecret"]);
//    Console.WriteLine(myToken.access_token);

//AdAppToken myToken = CsRestSharp_GetAzureTokenApplicationCertificate(
//    ConfigurationManager.AppSettings["TenantName"],
//    ConfigurationManager.AppSettings["ClientIdWithCert"],
//    ConfigurationManager.AppSettings["CertificateFilePath"],
//    ConfigurationManager.AppSettings["CertificateFilePw"],
//    ConfigurationManager.AppSettings["CertificateThumbprint"]);
//    Console.WriteLine(myToken.access_token);

//CsRestSharp_GetTeamDel();    
//CsRestSharp_CreateChannelDel();
//CsRestSharp_GetChannelDel(); 
//CsRestSharp_UpdateChannelDel();
//CsRestSharp_DeleteChannelDel();
//AdAppToken myToken = CsRestSharp_GetAzureTokenDelegation(
//    ConfigurationManager.AppSettings["TenantName"],
//    ConfigurationManager.AppSettings["ClientIdWithAccPw"],
//    ConfigurationManager.AppSettings["UserName"],
//    ConfigurationManager.AppSettings["UserPw"]);
//    Console.WriteLine(myToken.access_token);

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 002
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
//gavdcodeend 002

#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
