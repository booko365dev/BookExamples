using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Text;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static GraphServiceClient CsGraphMsal_GetGraphClientWithAccPw(
                                string TenantIdToConn, string ClientIdToConn,
                                string UserToConn, string PasswordToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    UsernamePasswordCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };
    UsernamePasswordCredential accPwCred = new(UserToConn, PasswordToConn,
                            TenantIdToConn, ClientIdToConn, clientOptionsCredential);
    
    GraphServiceClient graphClient = new(accPwCred, myScopes);

    return graphClient;
}
//gavdcodeend 001

//gavdcodebegin 002
static Tuple<string, string> CsGraphMsal_GetTokenWithAccPw(
                                string TenantIdToConn, string ClientIdToConn,
                                string UserToConn, string PasswordToConn)
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string[] myScopes = ["https://graph.microsoft.com/.default"];
    string myAuthority = $"https://login.microsoft.com/{TenantIdToConn}";

    IPublicClientApplication myApp = PublicClientApplicationBuilder
        .Create(ClientIdToConn)
        .WithAuthority(new Uri(myAuthority))
        .Build();

    try
    {
        AuthenticationResult myResult = myApp.AcquireTokenByUsernamePassword(myScopes, 
                    UserToConn, PasswordToConn).ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 002

//gavdcodebegin 003
static GraphServiceClient CsGraphMsal_GetGraphClientWithSecret(
                                string TenantIdToConn, string ClientIdToConn, 
                                string ClientSecret)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];
    
    ClientSecretCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };
    ClientSecretCredential clientSecretCred = new(TenantIdToConn, ClientIdToConn,
                            ClientSecret, clientOptionsCredential);
    
    GraphServiceClient graphClient = new(clientSecretCred, myScopes);

    return graphClient;
}
//gavdcodeend 003

//gavdcodebegin 004
static Tuple<string, string> CsGraphMsal_GetTokenWithSecret(
                                string TenantIdToConn, string ClientIdToConn,
                                string ClientSecret)
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string[] myScopes = ["https://graph.microsoft.com/.default"];
    string myAuthority = $"https://login.microsoftonline.com/{TenantIdToConn}";

    IConfidentialClientApplication myApp = ConfidentialClientApplicationBuilder
        .Create(ClientIdToConn)
        .WithAuthority(new Uri(myAuthority))
        .WithClientSecret(ClientSecret)
        .Build();

    try
    {
        AuthenticationResult myResult = myApp.AcquireTokenForClient(myScopes)
                    .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 004

//gavdcodebegin 005
static GraphServiceClient CsGraphMsal_GetGraphClientWithCertificateFile(
                                string TenantIdToConn, string ClientIdToConn, 
                                string CertificatePath, string CertificatePassword)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];
    
    X509Certificate2 myCertificate = new(CertificatePath, CertificatePassword);
    ClientCertificateCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };
    ClientCertificateCredential clientCertificateCred = new(TenantIdToConn, 
                                ClientIdToConn, myCertificate, clientOptionsCredential);
    
    GraphServiceClient graphClient = new(clientCertificateCred, myScopes);

    return graphClient;
}
//gavdcodeend 005

//gavdcodebegin 006
static Tuple<string, string> CsGraphMsal_GetTokenWithCertificateFile(
                                string TenantIdToConn, string ClientIdToConn,
                                string CertificatePath, string CertificatePassword)
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string[] myScopes = ["https://graph.microsoft.com/.default"];
    string myAuthority = $"https://login.microsoftonline.com/{TenantIdToConn}";

    // Load the myCertificate from a .pfx file
    X509Certificate2 myCertificate = new(CertificatePath, CertificatePassword);

    IConfidentialClientApplication myApp = ConfidentialClientApplicationBuilder
        .Create(ClientIdToConn)
        .WithAuthority(new Uri(myAuthority))
        .WithCertificate(myCertificate)
        .Build();

    try
    {
        AuthenticationResult myResult = myApp.AcquireTokenForClient(myScopes)
                    .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 006

//gavdcodebegin 007
static GraphServiceClient CsGraphMsal_GetGraphClientWithCertificateThumbprint(
                                string TenantIdToConn, string ClientIdToConn,
                                string CertificateThumbprint)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    X509Store myStore = new(StoreName.My, StoreLocation.CurrentUser);
    myStore.Open(OpenFlags.ReadOnly);
    X509Certificate2Collection certCollection = myStore.Certificates
                .Find(X509FindType.FindByThumbprint, CertificateThumbprint, false);
    X509Certificate2 myCertificate = certCollection[0];
    myStore.Close();

    ClientCertificateCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };
    ClientCertificateCredential clientCertificateCred = new(TenantIdToConn,
                                ClientIdToConn, myCertificate, clientOptionsCredential);

    GraphServiceClient graphClient = new(clientCertificateCred, myScopes);

    return graphClient;
}
//gavdcodeend 007

//gavdcodebegin 008
static Tuple<string, string> CsGraphMsal_GetTokenWithCertificateThumbprint(
                                string TenantIdToConn, string ClientIdToConn,
                                string CertificateThumbprint)
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string[] myScopes = ["https://graph.microsoft.com/.default"];
    string myAuthority = $"https://login.microsoftonline.com/{TenantIdToConn}";

    // Load the myCertificate from thumbprint
    X509Store myStore = new(StoreName.My, StoreLocation.CurrentUser);
    myStore.Open(OpenFlags.ReadOnly);
    X509Certificate2Collection certCollection = myStore.Certificates
                .Find(X509FindType.FindByThumbprint, CertificateThumbprint, false);
    X509Certificate2 myCertificate = certCollection[0];
    myStore.Close();

    IConfidentialClientApplication myApp = ConfidentialClientApplicationBuilder
        .Create(ClientIdToConn)
        .WithAuthority(new Uri(myAuthority))
        .WithCertificate(myCertificate)
        .Build();

    try
    {
        AuthenticationResult myResult = myApp.AcquireTokenForClient(myScopes)
                    .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 008

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 009
static void CsGraphRestApi_GetUsers_UsingGraphToken()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"]; 
    string myClientIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"]; 
    string myUserName = ConfigurationManager.AppSettings["UserName"]; 
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];
    string myClientIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
    string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCerFilePath = ConfigurationManager.AppSettings["CertificateFilePath"]; ;
    string myCerFilePw = ConfigurationManager.AppSettings["CertificateFilePw"]; ;
    string myCerThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"]; ;

    Tuple<string, string> myToken = CsGraphMsal_GetTokenWithAccPw(
                            myTenantId, myClientIdWithAccPw, myUserName, myUserPw);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithSecret(
    //                    myTenantId, myClientIdWithSecret, myClientSecret);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateFile(
    //                    myTenantId, myClientIdWithCer, myCerFilePath, myCerFilePw);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateThumbprint(
    //                    myTenantId, myClientIdWithCer, myCerThumbprint);

    if (myToken.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = "https://graph.microsoft.com/v1.0/users";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                "Authorization", "Bearer " + myToken.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                "Accept", "application/json"); // Output as JSON

        string myRequestBodyContent = myHttpClient.GetAsync(myEndpoint)
                                                  .ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        if (string.IsNullOrEmpty(myRequestBodyContent) == true)
        {
            Console.WriteLine("No users found");
        }
        else
        {
            Console.WriteLine(myRequestBodyContent);
        }
    }
    else
    {
        Console.WriteLine(myToken.Item2);  // Error retrieving the token
    }

}
//gavdcodeend 009

//gavdcodebegin 010
static void CsGraphSdk_GetUsers_UsingGraphClient()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];
    string myClientIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
    string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCerFilePath = ConfigurationManager.AppSettings["CertificateFilePath"]; ;
    string myCerFilePw = ConfigurationManager.AppSettings["CertificateFilePw"]; ;
    string myCerThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"]; ;

    GraphServiceClient myGraphClient = CsGraphMsal_GetGraphClientWithAccPw(
                            myTenantId, myClientIdWithAccPw, myUserName, myUserPw);
    //GraphServiceClient myGraphClient = CsGraphMsal_GetGraphClientWithSecret(
    //                        myTenantId, myClientIdWithSecret, myClientSecret);
    //GraphServiceClient myGraphClient = CsGraphMsal_GetGraphClientWithCertificateFile(
    //                        myTenantId, myClientIdWithCer, myCerFilePath, myCerFilePw);
    //GraphServiceClient myGraphClient = CsGraphMsal_GetGraphClientWithCertificateThumbprint(
    //                        myTenantId, myClientIdWithCer, myCerThumbprint);

    UserCollectionResponse myUsers = myGraphClient.Users.GetAsync().Result;

    foreach (User oneUser in myUsers!.Value!)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsGraphSdk_GetUsers_UsingGraphToken()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];
    string myClientIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
    string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCerFilePath = ConfigurationManager.AppSettings["CertificateFilePath"]; ;
    string myCerFilePw = ConfigurationManager.AppSettings["CertificateFilePw"]; ;
    string myCerThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"]; ;

    Tuple<string, string> myToken = CsGraphMsal_GetTokenWithAccPw(
                        myTenantId, myClientIdWithAccPw, myUserName, myUserPw);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithSecret(
    //                    myTenantId, myClientIdWithSecret, myClientSecret);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateFile(
    //                    myTenantId, myClientIdWithCer, myCerFilePath, myCerFilePw);
    //Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateThumbprint(
    //                    myTenantId, myClientIdWithCer, myCerThumbprint);

    HttpClient myHttpClient = new();
    myHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(
                        "Bearer", myToken.Item2);
    GraphServiceClient myGraphClient = new(myHttpClient);

    UserCollectionResponse myUsers = myGraphClient.Users.GetAsync().Result;

    foreach (User oneUser in myUsers!.Value!)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void Cs_GetJWTAssertionForAccPw()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    // Create the JWT header and payload
    DateTime parNow = DateTime.UtcNow;
    JwtSecurityTokenHandler jwtTokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Audience = $"https://login.microsoftonline.com/{myTenantId}/oauth2/v2.0/token",
        Issuer = myClientIdWithAccPw,
        Subject = new ClaimsIdentity(
            [
                new Claim(JwtRegisteredClaimNames.Sub, myUserName),
                new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
            ]),
        NotBefore = parNow,
        Expires = parNow.AddMinutes(10),
        SigningCredentials = new SigningCredentials(
            new SymmetricSecurityKey(Encoding.UTF8.GetBytes(myUserPw.PadRight(32, '0'))),
            SecurityAlgorithms.HmacSha256)
    };

    // Generate the JWT assertion
    SecurityToken jwtAssertion = jwtTokenHandler.CreateToken(tokenDescriptor);

    Console.WriteLine(jwtTokenHandler.WriteToken(jwtAssertion));
}
//gavdcodeend 012

//gavdcodebegin 013
static void Cs_GetJWTAssertionForSecret()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    // Create the JWT header and payload
    DateTime parNow = DateTime.UtcNow;
    JwtSecurityTokenHandler jwtTokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Audience = $"https://login.microsoftonline.com/{myTenantId}/oauth2/v2.0/token",
        Issuer = myClientIdWithSecret,
        Subject = new ClaimsIdentity(
            [
                new Claim(JwtRegisteredClaimNames.Sub, myClientIdWithSecret),
                new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
            ]),
        NotBefore = parNow,
        Expires = parNow.AddMinutes(10),
        SigningCredentials = new SigningCredentials(
                new SymmetricSecurityKey(Encoding.UTF8.GetBytes(myClientSecret)),
                SecurityAlgorithms.HmacSha256)
    };

    // Generate the JWT assertion
    SecurityToken jwtAssertion = jwtTokenHandler.CreateToken(tokenDescriptor);

    Console.WriteLine(jwtTokenHandler.WriteToken(jwtAssertion));
}
//gavdcodeend 013

//gavdcodebegin 014
static void Cs_GetJWTAssertionForCertificateFile()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCerFilePath = ConfigurationManager.AppSettings["CertificateFilePath"]; ;
    string myCerFilePw = ConfigurationManager.AppSettings["CertificateFilePw"]; ;

    // Load the myCertificate from a .pfx file
    X509Certificate2 myCert = new(myCerFilePath, myCerFilePw);

    // Create the JWT header and payload
    X509SigningCredentials signingCredentials = new(myCert);
    // Create the JWT header and payload
    DateTime parNow = DateTime.UtcNow;
    JwtSecurityTokenHandler jwtTokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Audience = $"https://login.microsoftonline.com/{myTenantId}/oauth2/v2.0/token",
        Issuer = myClientIdWithCer,
        Subject = new ClaimsIdentity(
        [
            new Claim(JwtRegisteredClaimNames.Sub, myClientIdWithCer),
            new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
        ]),
        NotBefore = parNow,
        Expires = parNow.AddMinutes(10),
        SigningCredentials = signingCredentials
    };

    // Generate the JWT assertion
    SecurityToken jwtAssertion = jwtTokenHandler.CreateToken(tokenDescriptor);
    
    Console.WriteLine(jwtTokenHandler.WriteToken(jwtAssertion));
}
//gavdcodeend 014

//gavdcodebegin 015
static void Cs_GetJWTAssertionForCertificateThumbprint()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCerThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"]; ;

    // Load the myCertificate from thumbprint
    X509Store myStore = new(StoreName.My, StoreLocation.CurrentUser);
    myStore.Open(OpenFlags.ReadOnly);
    X509Certificate2Collection certCollection = myStore.Certificates
                .Find(X509FindType.FindByThumbprint, myCerThumbprint, false);
    if (certCollection.Count == 0)
    {
        throw new Exception("Certificate not found");
    }
    X509Certificate2 myCert = certCollection[0];
    myStore.Close();

    // Create the JWT header and payload
    X509SigningCredentials signingCredentials = new(myCert);
    DateTime parNow = DateTime.UtcNow;
    JwtSecurityTokenHandler jwtTokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Audience = $"https://login.microsoftonline.com/{myTenantId}/oauth2/v2.0/token",
        Issuer = myClientIdWithCer,
        Subject = new ClaimsIdentity(
        [
            new Claim(JwtRegisteredClaimNames.Sub, myClientIdWithCer),
            new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
        ]),
        NotBefore = parNow,
        Expires = parNow.AddMinutes(10),
        SigningCredentials = signingCredentials
    };

    // Generate the JWT assertion
    SecurityToken jwtAssertion = jwtTokenHandler.CreateToken(tokenDescriptor);

    Console.WriteLine(jwtTokenHandler.WriteToken(jwtAssertion));
}
//gavdcodeend 015

//gavdcodebegin 016
static void Cs_GetTokenFromJWTAssertion()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myAudience = $"https://login.microsoftonline.com/{myTenantId}/oauth2/v2.0/token";
    string myScope = "https://graph.microsoft.com/.default";
    string myGrantType = "client_credentials";  // or "password"
    string myAssertionType = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";
    string myAssertion = "eyJhbGciOiJSUzI1N...wk4DdgTvB9akbIJp-g";

    HttpClient myClient = new();
    HttpRequestMessage myRequest = new(HttpMethod.Post, myAudience);
    List<KeyValuePair<string, string>> contentCollection =
    [
        new("grant_type", myGrantType),
        new("client_id", myClientId),
        new("client_assertion_type", myAssertionType),
        new("client_assertion", myAssertion),
        new("scope", myScope),
    ];
    FormUrlEncodedContent myContent = new(contentCollection);
    myRequest.Content = myContent;

    HttpResponseMessage myResponse = myClient.SendAsync(myRequest).Result;
    myResponse.EnsureSuccessStatusCode();
    string myTokenJSON = myResponse.Content.ReadAsStringAsync().Result;

    JObject tokenObj = JObject.Parse(myTokenJSON);
    string accessToken = tokenObj["access_token"].ToString();

    Console.WriteLine("Full token as JSON: " + myTokenJSON);
    Console.WriteLine("Access token: " + accessToken);
}
//gavdcodeend 016

//gavdcodebegin 017
static void Cs_UseTokenFromJWTAssertion()
{
    string myQuery = "https://graph.microsoft.com/v1.0/users";
    string myAccessToken = "eyJ0eXAiOiJKV1QiLCJu...k1ZGQBuqQGuGK7zQQ";

    var myClient = new HttpClient();
    var myRequest = new HttpRequestMessage(HttpMethod.Get, myQuery);
    myRequest.Headers.Add("Authorization", myAccessToken);
    var myResponse = myClient.SendAsync(myRequest).Result;
    myResponse.EnsureSuccessStatusCode();

    string myResult = myResponse.Content.ReadAsStringAsync().Result;

    Console.WriteLine("All users as JSON: " + myResult);

    JObject jsonResult = JObject.Parse(myResult);
    var allUsers = jsonResult["value"];

    if (allUsers != null)
    {
        foreach (JToken oneUser in allUsers)
        {
            Console.WriteLine(oneUser["displayName"]);
        }
    }
    else
    {
        Console.WriteLine("No users found");
    }
}
//gavdcodeend 017

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 017 ***

string myTenantId = ConfigurationManager.AppSettings["TenantName"];
string myClientIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
string myUserName = ConfigurationManager.AppSettings["UserName"];
string myUserPw = ConfigurationManager.AppSettings["UserPw"];
string myClientIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
string myClientIdWithCer = ConfigurationManager.AppSettings["ClientIdWithCert"];
string myCerFilePath = ConfigurationManager.AppSettings["CertificateFilePath"]; ;
string myCerFilePw = ConfigurationManager.AppSettings["CertificateFilePw"]; ;
string myCerThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"]; ;

//CsGraphRestApi_GetUsers_UsingGraphToken();
//CsGraphSdk_GetUsers_UsingGraphClient();
//CsGraphSdk_GetUsers_UsingGraphToken();

// Test the token retrieval
//Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithAccPw(
//                        myTenantId, myClientIdWithAccPw, myUserName, myUserPw);
//Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithSecret(
//                    myTenantId, myClientIdWithSecret, myClientSecret);
//Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateFile(
//                    myTenantId, myClientIdWithCer, myCerFilePath, myCerFilePw);
//Tuple<string, string> myTokenJSON = CsGraphMsal_GetTokenWithCertificateThumbprint(
//                    myTenantId, myClientIdWithCer, myCerThumbprint);
//Console.WriteLine(myTokenJSON.Item2);

//****************************************************
//*** Auxiliary routines
//Cs_GetJWTAssertionForAccPw();
//Cs_GetJWTAssertionForSecret();
//Cs_GetJWTAssertionForCertificateFile();
//Cs_GetJWTAssertionForCertificateThumbprint();
//Cs_GetTokenFromJWTAssertion();
//Cs_UseTokenFromJWTAssertion();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
