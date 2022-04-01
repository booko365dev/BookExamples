using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Configuration;
using System.Security;
using System.Security.Cryptography.X509Certificates;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 05
static GraphServiceClient GetGraphClientWithInteraction(string ClientId)
{
    string[] myScopes = new string[]
    {
                "https://graph.microsoft.com/.default"
    };

    InteractiveBrowserCredentialOptions clientInteractCredential =
        new InteractiveBrowserCredentialOptions()
        {
            ClientId = ClientId
        };
    InteractiveBrowserCredential interactBrowserCredential =
        new InteractiveBrowserCredential(clientInteractCredential);
    GraphServiceClient graphClient =
        new GraphServiceClient(interactBrowserCredential, myScopes);

    return graphClient;
}
//gavdcodeend 05

//gavdcodebegin 06
static GraphServiceClient GetGraphClientWithAccPw(
            string TenantId, string ClientId, string Account, string Password)
{
    string[] myScopes = new string[]
    {
                "https://graph.microsoft.com/.default"
    };

    UsernamePasswordCredential clientAccPwCredential =
        new UsernamePasswordCredential(Account, Password, TenantId, ClientId);
    GraphServiceClient graphClient =
        new GraphServiceClient(clientAccPwCredential, myScopes);

    return graphClient;
}
//gavdcodeend 06

//gavdcodebegin 07
static GraphServiceClient GetGraphClientWithSecret(
                            string TenantId, string ClientId, string ClientSecret)
{
    string[] myScopes = new string[]
    {
                "https://graph.microsoft.com/.default"
    };

    ClientSecretCredential clientSecretCredential =
        new ClientSecretCredential(TenantId, ClientId, ClientSecret);
    GraphServiceClient graphClient =
        new GraphServiceClient(clientSecretCredential, myScopes);

    return graphClient;
}
//gavdcodeend 07

//gavdcodebegin 08
static GraphServiceClient GetGraphClientWithCertificat(
                   string TenantId, string ClientId, string CertificateThumbprint,
                   StoreName CertStoreName, StoreLocation CertStoreLocation)
{
    X509Certificate2 myCert = GetCertificateByThumbprint(
                   StoreName.My, StoreLocation.CurrentUser, CertificateThumbprint);

    string[] myScopes = new string[]
    {
                "https://graph.microsoft.com/.default"
    };

    ClientCertificateCredential clientCertificateCredential =
        new ClientCertificateCredential(TenantId, ClientId, myCert);
    GraphServiceClient graphClient = new GraphServiceClient(
                                           clientCertificateCredential, myScopes);

    return graphClient;
}
//gavdcodeend 08

//gavdcodebegin 09
static void GetTokenWithInteraction(string TenantId, string ClientId)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IPublicClientApplication myPubClientApp = PublicClientApplicationBuilder
            .Create(ClientId)
            .WithRedirectUri("http://localhost")
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = new List<string>()
                {
                    "https://management.azure.com/.default"
                };

    AuthenticationResult myToken = myPubClientApp
            .AcquireTokenInteractive(myScopes)
            .ExecuteAsync()
            .Result;

    Console.WriteLine("Token for   - " + myToken.Account.Username);
    Console.WriteLine("Token value - " + myToken.AccessToken);
}
//gavdcodeend 09

//gavdcodebegin 10
static void GetTokenWithAccPw(
                        string TenantId, string ClientId, string Account, string Password)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IPublicClientApplication myPubClientApp = PublicClientApplicationBuilder
            .Create(ClientId)
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = new List<string>()
                {
                    "https://management.azure.com/.default"
                };

    SecureString usrPw = new SecureString();
    foreach (char oneChar in Password)
        usrPw.AppendChar(oneChar);

    AuthenticationResult myToken = myPubClientApp
            .AcquireTokenByUsernamePassword(myScopes, Account, usrPw)
            .ExecuteAsync()
            .Result;

    Console.WriteLine("Token for   - " + myToken.Account.Username);
    Console.WriteLine("Token value - " + myToken.AccessToken);
}
//gavdcodeend 10

//gavdcodebegin 11
static void GetTokenWithSecret(string TenantId, string ClientId, string ClientSecret)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IConfidentialClientApplication myConfClientApp =
        ConfidentialClientApplicationBuilder
            .Create(ClientId)
            .WithClientSecret(ClientSecret)
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = new List<string>()
            {
                "https://management.azure.com/.default"
            };

    AuthenticationResult myToken = myConfClientApp
            .AcquireTokenForClient(myScopes)
            .ExecuteAsync()
            .Result;

    Console.WriteLine("Token value - " + myToken.AccessToken);
}
//gavdcodeend 11

//gavdcodebegin 12
static void GetTokenWithCertificate(
                        string TenantId, string ClientId, string CertificateThumbprint)
{
    string authorityEndpoint = "https://login.microsoftonline.com/" + TenantId;

    IConfidentialClientApplication myConfClientApp =
            ConfidentialClientApplicationBuilder
            .Create(ClientId)
            .WithCertificate(GetCertificateByThumbprint(
                  StoreName.My, StoreLocation.CurrentUser, CertificateThumbprint))
            //.WithCertificate(GetCertificateByName(
            //      StoreName.My, StoreLocation.CurrentUser, certName))
            .WithAuthority(new Uri(authorityEndpoint))
            .Build();

    List<string> myScopes = new List<string>()
            {
                "https://management.azure.com/.default"
            };

    AuthenticationResult myToken = myConfClientApp
            .AcquireTokenForClient(myScopes)
            .ExecuteAsync()
            .Result;

    Console.WriteLine("Token value - " + myToken.AccessToken);
}
//gavdcodeend 12

//gavdcodebegin 13
static X509Certificate2 GetCertificateByThumbprint(
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

static X509Certificate2 GetCertificateByName(
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
//gavdcodeend 13

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static void GetGrQueryWithInteraction()
{
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];

    GraphServiceClient myGraphClient = GetGraphClientWithInteraction(myClientId);
    Site mySite = (Site)myGraphClient
            .Sites["077e9977-65c5-4acf-947e-63552e7573b4"]
            .Request()
            .GetAsync().Result;

    Console.WriteLine(mySite.Name);
}
//gavdcodeend 01

//gavdcodebegin 02
static void GetGrQueryWithAccPw()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    GraphServiceClient myGraphClient =
                    GetGraphClientWithAccPw(myTenantId, myClientId, myUserName, myUserPw);
    SiteListsCollectionPage myLists = (SiteListsCollectionPage)myGraphClient
            .Sites["077e9977-65c5-4acf-947e-63552e7573b4"]
            .Lists
            .Request()
            .GetAsync().Result;

    foreach (List oneList in myLists)
    {
        Console.WriteLine(oneList.Name);
    }
}
//gavdcodeend 02

//gavdcodebegin 03
static void GetGrQueryWithSecret()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
                    GetGraphClientWithSecret(myTenantId, myClientId, myClientSecret);
    ListItemsCollectionPage myItems = (ListItemsCollectionPage)myGraphClient
            .Sites["077e9977-65c5-4acf-947e-63552e7573b4"]
            .Lists["Documents"]
            .Items
            .Request()
            .GetAsync().Result;

    foreach (ListItem oneItem in myItems)
    {
        Console.WriteLine(oneItem.WebUrl);
    }
}
//gavdcodeend 03

//gavdcodebegin 04
static void GetGrQueryWithCertificate()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string myCertThumbprintt = ConfigurationManager.AppSettings["CertificateThumbprint"];

    GraphServiceClient myGraphClient =
          GetGraphClientWithCertificat(myTenantId, myClientId, myCertThumbprintt,
                                        StoreName.My, StoreLocation.CurrentUser);
    User myUser = (User)myGraphClient
            .Users[ConfigurationManager.AppSettings["UserName"]]
            .Request()
            .GetAsync().Result;

    Console.WriteLine(myUser.DisplayName);
}
//gavdcodeend 04

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

string myTenantId = ConfigurationManager.AppSettings["TenantName"];
string myClIdWithSecret = ConfigurationManager.AppSettings["ClientIdWithSecret"];
string myClSecret = ConfigurationManager.AppSettings["ClientSecret"];
string myClIdWithAccPw = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
string myUserName = ConfigurationManager.AppSettings["UserName"];
string myUserPw = ConfigurationManager.AppSettings["UserPw"];
string myClIdWithCert = ConfigurationManager.AppSettings["ClientIdWithCert"];
string myCertThumbpr = ConfigurationManager.AppSettings["CertificateThumbprint"];

//GetGrQueryWithInteraction();
//GetGrQueryWithAccPw();
//GetGrQueryWithSecret();
//GetGrQueryWithCertificate();

//GetTokenWithInteraction(myTenantId, myClIdWithAccPw);
//GetTokenWithAccPw(myTenantId, myClIdWithAccPw, myUserName, myUserPw);
//GetTokenWithSecret(myTenantId, myClIdWithSecret, myClSecret);
//GetTokenWithCertificate(myTenantId, myClIdWithCert, myCertThumbpr);

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
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
