using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using System.Configuration;
using System.Security.Cryptography.X509Certificates;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//----------------- Get Clients

//gavdcodebegin 005
static GraphServiceClient CsGraphSdk_LoginWithInteraction(
                                string TenantIdToConn, string ClientIdToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    InteractiveBrowserCredentialOptions clientOptionsCredential = new()
    {
        TenantId = TenantIdToConn,
        ClientId = ClientIdToConn,
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        RedirectUri = new Uri("http://localhost")
        // RedirectUri MUST be http://localhost or http://localhost:PORT
    };
    InteractiveBrowserCredential interactBrowserCredential =
        new(clientOptionsCredential);
    GraphServiceClient graphClient = new(interactBrowserCredential, myScopes);

    return graphClient;
}
//gavdcodeend 005

//gavdcodebegin 014
static GraphServiceClient CsGraphSdk_LoginWithDeviceCode(
                                string TenantIdToConn, string ClientIdToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    DeviceCodeCredentialOptions clientOptionsCredential = new()
    {
        TenantId = TenantIdToConn,
        ClientId = ClientIdToConn,
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        DeviceCodeCallback = (code, cancellation) =>
        {
            Console.WriteLine(code.Message);
            return Task.FromResult(0);
        }
    };
    DeviceCodeCredential deviceCodeCredential = new(clientOptionsCredential);
    GraphServiceClient graphClient = new(deviceCodeCredential, myScopes);

    return graphClient;
}
//gavdcodeend 014

//gavdcodebegin 006
static GraphServiceClient CsGraphSdk_LoginWithAccPw(
                                string TenantIdToConn, string ClientIdToConn,
                                string UserToConn, string PasswordToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    UsernamePasswordCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    UsernamePasswordCredential accPwCredential =
                    new(UserToConn, PasswordToConn, TenantIdToConn,
                        ClientIdToConn, clientOptionsCredential);
    GraphServiceClient graphClient = new(accPwCredential, myScopes);

    return graphClient;
}
//gavdcodeend 006

//gavdcodebegin 017
static Microsoft.Graph.Beta.GraphServiceClient CsGraphSdk_LoginWithAccPwBeta(
                                string TenantIdToConn, string ClientIdToConn,
                                string UserToConn, string PasswordToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    UsernamePasswordCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    UsernamePasswordCredential accPwCredential =
                    new(UserToConn, PasswordToConn, TenantIdToConn,
                        ClientIdToConn, clientOptionsCredential);
    Microsoft.Graph.Beta.GraphServiceClient graphClient = new(accPwCredential, myScopes);

    return graphClient;
}
//gavdcodeend 017

//gavdcodebegin 007
static GraphServiceClient CsGraphSdk_LoginWithSecret(
                                string TenantIdToConn, string ClientIdToConn,
                                string ClientSecretToConn)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    ClientSecretCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
    };

    ClientSecretCredential secretCredential =
                    new(TenantIdToConn, ClientIdToConn, ClientSecretToConn,
                        clientOptionsCredential);
    GraphServiceClient graphClient = new(secretCredential, myScopes);

    return graphClient;
}
//gavdcodeend 007

//gavdcodebegin 008
static GraphServiceClient CsGraphSdk_LoginWithCertificateThumbprint(
                                string TenantIdToConn, string ClientIdToConn,
                                string CertificateThumbprintToConn,
                                StoreName CertStoreName, StoreLocation CertStoreLocation)
{
    X509Certificate2 myCert = CsGraphSdk_GetCertificateByThumbprint(
                    StoreName.My, StoreLocation.CurrentUser,
                    CertificateThumbprintToConn);

    string[] myScopes = ["https://graph.microsoft.com/.default"];

    ClientCertificateCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
    };

    ClientCertificateCredential certificateCredential =
                new(TenantIdToConn, ClientIdToConn, myCert, clientOptionsCredential);
    GraphServiceClient graphClient = new(certificateCredential, myScopes);

    return graphClient;
}
//gavdcodeend 008

//gavdcodebegin 020
static GraphServiceClient CsGraphSdk_LoginWithCertificateFile(
                                string TenantIdToConn, string ClientIdToConn,
                                string CertPathToConn, string CertPwToConn)
{
    X509Certificate2 myCert = new(CertPathToConn, CertPwToConn);

    string[] myScopes = ["https://graph.microsoft.com/.default"];

    ClientCertificateCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
    };

    ClientCertificateCredential certificateCredential =
                new(TenantIdToConn, ClientIdToConn, myCert, clientOptionsCredential);
    GraphServiceClient graphClient = new(certificateCredential, myScopes);

    return graphClient;
}
//gavdcodeend 020

//gavdcodebegin 022
static GraphServiceClient CsGraphSdk_LoginWithToken(string AccessToken)
{
    string[] myScopes = ["https://graph.microsoft.com/.default"];

    TokenCredential tokenCredential = new AccessTokenCredential(AccessToken);
    // AccessTokenCredential is a custom class, see below
    GraphServiceClient graphClient = new(tokenCredential, myScopes);

    return graphClient;
}
//gavdcodeend 022

//----------------- Get Tokens

//gavdcodebegin 009
static string CsGraphSdk_GetTokenWithInteraction(string TenantId, string ClientId)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IPublicClientApplication myPubClientApp = PublicClientApplicationBuilder
                        .Create(ClientId)
                        .WithRedirectUri("http://localhost")
                        .WithAuthority(new Uri(authorityEndpoint))
                        .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myPubClientApp
                        .AcquireTokenInteractive(myScopes)
                        .ExecuteAsync()
                        .Result;

    //Console.WriteLine("Token for   - " + myToken.Account.Username);
    //Console.WriteLine("Token value - " + myToken.AccessToken);
    return myToken.AccessToken;
}
//gavdcodeend 009

//gavdcodebegin 016
static string CsGraphSdk_GetTokenWithDeviceCode(string TenantId, string ClientId)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IPublicClientApplication myPubClientApp = PublicClientApplicationBuilder
                        .Create(ClientId)
                        .WithAuthority(new Uri(authorityEndpoint))
                        .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myPubClientApp
                        .AcquireTokenWithDeviceCode(myScopes, code =>
                            {
                                Console.WriteLine(code.Message);
                                return Task.FromResult(0);
                            })
                        .ExecuteAsync()
                        .Result;

    //Console.WriteLine("Token for   - " + myToken.Account.Username);
    //Console.WriteLine("Token value - " + myToken.AccessToken);
    return myToken.AccessToken;
}
//gavdcodeend 016

//gavdcodebegin 010
static string CsGraphSdk_GetTokenWithAccPw(
                      string TenantId, string ClientId, string Account, string Password)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IPublicClientApplication myPubClientApp = PublicClientApplicationBuilder
                        .Create(ClientId)
                        .WithAuthority(new Uri(authorityEndpoint))
                        .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myPubClientApp
                        .AcquireTokenByUsernamePassword(myScopes, Account, Password)
                        .ExecuteAsync()
                        .Result;

    Console.WriteLine("Token for   - " + myToken.Account.Username);
    Console.WriteLine("Token value - " + myToken.AccessToken);

    //Console.WriteLine("Token value - " + myToken.AccessToken);
    return myToken.AccessToken;
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsGraphSdk_GetTokenWithSecret(
                        string TenantId, string ClientId, string ClientSecret)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IConfidentialClientApplication myConfClientApp = ConfidentialClientApplicationBuilder
                        .Create(ClientId)
                        .WithClientSecret(ClientSecret)
                        .WithAuthority(new Uri(authorityEndpoint))
                        .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myConfClientApp
                        .AcquireTokenForClient(myScopes)
                        .ExecuteAsync()
                        .Result;

    Console.WriteLine("Token value - " + myToken.AccessToken);
}
//gavdcodeend 011

//gavdcodebegin 012
static string CsGraphSdk_GetTokenWithCertificateThumbprint(
                        string TenantId, string ClientId, string CertificateThumbprint)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IConfidentialClientApplication myConfClientApp = ConfidentialClientApplicationBuilder
            .Create(ClientId)
            .WithCertificate(CsGraphSdk_GetCertificateByThumbprint(
                  StoreName.My, StoreLocation.CurrentUser, CertificateThumbprint))
            //.WithCertificate(CsGraphSdk_GetCertificateByName(
            //      StoreName.My, StoreLocation.CurrentUser, certName))
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myConfClientApp
                        .AcquireTokenForClient(myScopes)
                        .ExecuteAsync()
                        .Result;

    //Console.WriteLine("Token value - " + myToken.AccessToken);
    return myToken.AccessToken;
}
//gavdcodeend 012

//gavdcodebegin 013
static X509Certificate2 CsGraphSdk_GetCertificateByThumbprint(
                        StoreName storeName, StoreLocation storeLoc, String thumbprint)
{
    X509Store myStore = new X509Store(storeName, storeLoc);
    X509Certificate2 myCertificat = null;

    myStore.Open(OpenFlags.ReadOnly);

    // Also for self signed, set the valid parameter to 'false'
    X509Certificate2Collection allCertificates = myStore.Certificates.Find(
                                X509FindType.FindByThumbprint, thumbprint, false);
    if (allCertificates.Count > 0)
    {
        myCertificat = allCertificates[0];
    };
    myStore.Close();

    return myCertificat;
}

static X509Certificate2 CsGraphSdk_GetCertificateByName(
                        StoreName storeName, StoreLocation storeLoc, String subjectName)
{
    X509Store myStore = new X509Store(storeName, storeLoc);
    X509Certificate2 myCertificate = null;

    myStore.Open(OpenFlags.ReadOnly);

    // Also for self signed, set the valid parameter to 'false'
    X509Certificate2Collection allCertificates = myStore.Certificates.Find(
                              X509FindType.FindBySubjectName, subjectName, false);
    if (allCertificates.Count > 0)
    {
        myCertificate = allCertificates[0];
    };
    myStore.Close();

    return myCertificate;
}
//gavdcodeend 013

//gavdcodebegin 019
static string CsGraphSdk_GetTokenWithCertificateFile(
        string TenantId, string ClientId, string CertificatePath, string CertificatePw)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    // Load the myCertificate from a .pfx file
    X509Certificate2 myCertificate = new(CertificatePath, CertificatePw);

    IConfidentialClientApplication myConfClientApp = ConfidentialClientApplicationBuilder
            .Create(ClientId)
            .WithCertificate(myCertificate)
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = ["https://graph.microsoft.com/.default"];
    //List<string> myScopes = ["https://management.azure.com/.default"];

    AuthenticationResult myToken = myConfClientApp
                        .AcquireTokenForClient(myScopes)
                        .ExecuteAsync()
                        .Result;

    //Console.WriteLine("Token value - " + myToken.AccessToken);
    return myToken.AccessToken;
}
//gavdcodeend 019

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsGraphSdk_GetQueryWithInteraction()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];

    GraphServiceClient myGraphClient = CsGraphSdk_LoginWithInteraction(
                                                myTenantId, myClientId);

    Site mySite = myGraphClient
                        .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                        .GetAsync().Result;

    Console.WriteLine(mySite.Name);
}
//gavdcodeend 001

//gavdcodebegin 015
static void CsGraphSdk_GetQueryWithDeviceCode()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];

    GraphServiceClient myGraphClient = CsGraphSdk_LoginWithDeviceCode(
                                                myTenantId, myClientId);

    var myTeams = myGraphClient.Teams.GetAsync().Result;
    foreach (var oneTeam in myTeams.Value)
    {
        Console.WriteLine(oneTeam.DisplayName);
    }
}
//gavdcodeend 015

//gavdcodebegin 002
static void CsGraphSdk_GetQueryWithAccPw()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    GraphServiceClient myGraphClient =
       CsGraphSdk_LoginWithAccPw(myTenantId, myClientId, myUserName, myUserPw);

    ListCollectionResponse myLists = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists
            .GetAsync().Result;

    foreach (List oneList in myLists.Value)
    {
        Console.WriteLine(oneList.Name);
    }
}
//gavdcodeend 002

//gavdcodebegin 018
static void CsGraphSdk_GetQueryWithAccPwBeta()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    Microsoft.Graph.Beta.GraphServiceClient myGraphClient =
       CsGraphSdk_LoginWithAccPwBeta(myTenantId, myClientId, myUserName, myUserPw);

    var myUsers = myGraphClient.Users.GetAsync().Result;
    foreach (var oneUser in myUsers.Value)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 018

//gavdcodebegin 003
static void CsGraphSdk_GetQueryWithSecret()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ListItemCollectionResponse myItems = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["Documents"]
            .Items
            .GetAsync().Result;

    foreach (ListItem oneItem in myItems.Value)
    {
        Console.WriteLine(oneItem.WebUrl);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsGraphSdk_GetQueryWithCertificateThumbprint()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCertThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"];

    GraphServiceClient myGraphClient =
          CsGraphSdk_LoginWithCertificateThumbprint(myTenantId, myClientId, myCertThumbprint,
                                        StoreName.My, StoreLocation.CurrentUser);

    UserCollectionResponse myUsers = myGraphClient
                .Users
                .GetAsync().Result;

    foreach (User oneUser in myUsers.Value)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 004

//gavdcodebegin 021
static void CsGraphSdk_GetQueryWithCertificateFile()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCertFilePath = ConfigurationManager.AppSettings["CertificateFilePath"];
    string myCertFilePw = ConfigurationManager.AppSettings["CertificateFilePw"];

    GraphServiceClient myGraphClient =
          CsGraphSdk_LoginWithCertificateFile(myTenantId, myClientId, myCertFilePath,
                                        myCertFilePw);

    UserCollectionResponse myUsers = myGraphClient
                .Users
                .GetAsync().Result;

    foreach (User oneUser in myUsers.Value)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 021

//gavdcodebegin 024
static void CsGraphSdk_GetQueryWithToken()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCertFilePath = ConfigurationManager.AppSettings["CertificateFilePath"];
    string myCertFilePw = ConfigurationManager.AppSettings["CertificateFilePw"];

    string myToken = CsGraphSdk_GetTokenWithCertificateFile(myTenantId, myClientId, 
                                                        myCertFilePath, myCertFilePw);
    GraphServiceClient myGraphClient = CsGraphSdk_LoginWithToken(myToken);

    UserCollectionResponse myUsers = myGraphClient
                .Users
                .GetAsync().Result;

    foreach (User oneUser in myUsers.Value)
    {
        Console.WriteLine(oneUser.DisplayName);
    }
}
//gavdcodeend 024

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 024 ***

string myTenantId = ConfigurationManager.AppSettings["TenantName"];
string myClIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
string myClSecret = ConfigurationManager.AppSettings["ClientSecret"];
string myClIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
string myUserName = ConfigurationManager.AppSettings["UserName"];
string myUserPw = ConfigurationManager.AppSettings["UserPw"];
string myClIdWithCert = ConfigurationManager.AppSettings["ClientIdWithCert"];
string myCertThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"];
string myCertFilePath = ConfigurationManager.AppSettings["CertificateFilePath"];
string myCertFilePw = ConfigurationManager.AppSettings["CertificateFilePw"];

//CsGraphSdk_GetQueryWithInteraction();
//CsGraphSdk_GetQueryWithDeviceCode();
//CsGraphSdk_GetQueryWithAccPw();
//CsGraphSdk_GetQueryWithAccPwBeta();
//CsGraphSdk_GetQueryWithSecret();
//CsGraphSdk_GetQueryWithCertificateThumbprint();
//CsGraphSdk_GetQueryWithCertificateFile();
//CsGraphSdk_GetQueryWithToken();

//CsGraphSdk_GetTokenWithInteraction(myTenantId, myClIdWithAccPw);
//CsGraphSdk_GetTokenWithDeviceCode(myTenantId, myClIdWithAccPw);
//CsGraphSdk_GetTokenWithAccPw(myTenantId, myClIdWithAccPw, myUserName, myUserPw);
//CsGraphSdk_GetTokenWithSecret(myTenantId, myClIdWithSecret, myClSecret);
//CsGraphSdk_GetTokenWithCertificateThumbprint(myTenantId, myClIdWithCert, myCertThumbprint);
//CsGraphSdk_GetTokenWithCertificateFile(myTenantId, myClIdWithCert, myCertFilePath, myCertFilePw);

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


//gavdcodebegin 023
class AccessTokenCredential : TokenCredential
{
    private readonly string _accessToken;

    public AccessTokenCredential(string accessToken)
    {
        _accessToken = accessToken;
    }

    public override AccessToken GetToken(TokenRequestContext requestContext, 
                                         CancellationToken cancellationToken)
    {
        return new AccessToken(_accessToken, DateTimeOffset.MaxValue);
    }

    public override ValueTask<AccessToken> GetTokenAsync(
                                        TokenRequestContext requestContext, 
                                        CancellationToken cancellationToken)
    {
        return new ValueTask<AccessToken>(new AccessToken(_accessToken, 
                                                        DateTimeOffset.MaxValue));
    }
}
//gavdcodeend 023


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
//---------------------------------------------------------------------------------------

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Legacy routines ***--------------------------
//---------------------------------------------------------------------------------------

#region Recipes based on the deprecated Microsoft.Graph.Auth --> Do not use anymore
// Do not implement - routine based on deprecated Microsoft.Graph.Auth
//static void GetGrQueryApp(string TenantId, string ClientId, string ClientSecret)
//{
//    GraphServiceClient graphClient = GetGraphClientApp(
//                                            TenantId, ClientId, ClientSecret);

//    User myMe = graphClient.Users[ConfigurationManager.AppSettings["UserName"]]
//        .Request()
//        .GetAsync().Result;
//    Console.WriteLine(myMe.DisplayName);

//    var myMessages = graphClient.Users[ConfigurationManager.AppSettings["UserName"]]
//        .Messages.Request()
//        .GetAsync().Result;
//    Console.WriteLine(myMessages.Count.ToString());
//}

// Do not implement - routine based on deprecated Microsoft.Graph.Auth
//static void GetGrQueryDel(string TenantId, string ClientId)
//{
//    GraphServiceClient graphClient = GetGraphClientDel(TenantId, ClientId);

//    var securePassword = new SecureString();
//    foreach (var chr in ConfigurationManager.AppSettings["UserPw"])
//    { securePassword.AppendChar(chr); }

//    User myMe = graphClient.Me.Request()
//        .WithUsernamePassword(ConfigurationManager.AppSettings
//                                                ["UserName"], securePassword)
//        .GetAsync().Result;
//    Console.WriteLine(myMe.DisplayName);

//    var myMessages = graphClient.Me.Messages.Request()
//        .WithUsernamePassword(ConfigurationManager.AppSettings
//                                                ["UserName"], securePassword)
//        .GetAsync().Result;
//    Console.WriteLine(myMessages.Count.ToString());
//}

// Do not implement - routine based on deprecated Microsoft.Graph.Auth
//static GraphServiceClient GetGraphClientApp(string TenantId, string ClientId, 
//                                                            string ClientSecret)
//{
//    IConfidentialClientApplication clientApplication = 
//        ConfidentialClientApplicationBuilder
//            .Create(ClientId)
//            .WithTenantId(TenantId)
//            .WithClientSecret(ClientSecret)
//            .Build();

//    ClientCredentialProvider authenticationProvider = 
//                            new ClientCredentialProvider(clientApplication);
//    GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);

//    return graphClient;
//}

// Do not implement - routine based on deprecated Microsoft.Graph.Auth
//static GraphServiceClient GetGraphClientDel(string TenantId, string ClientId)
//{
//    IPublicClientApplication clientApplication = PublicClientApplicationBuilder
//        .Create(ClientId)
//        .WithTenantId(TenantId)
//        //.WithRedirectUri("http://localhost") // Only if redirect in the App Reg.
//        .Build();

//    UsernamePasswordProvider authenticationProvider = 
//                            new UsernamePasswordProvider(clientApplication);
//    GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);

//    return graphClient;
//} 
#endregion
