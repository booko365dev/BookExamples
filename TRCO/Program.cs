using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Applications.Item.AddPassword;
using Microsoft.Graph.Applications.Item.RemovePassword;
using Microsoft.Graph.Models;
using Microsoft.IdentityModel.Tokens;
using System.Configuration;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used
#pragma warning disable CA1416 // Validate platform compatibility


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------
static GraphServiceClient CsEntraGraphCsSdk_LoginWithSecret()
{
    string TenantIdToConn = ConfigurationManager.AppSettings["TenantName"];
    string ClientIdToConn = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string ClientSecretToConn = ConfigurationManager.AppSettings["ClientSecret"];

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

// Routines that can be used in Azure Functions with Managed Identities
//gavdcodebegin 016
static GraphServiceClient CsEntraGraphCsSdk_LoginWithSecret_ForAzFuncts(
                                                       string clientId, string tenantId,
                                                       string clientSecret)
{
    ClientSecretCredential authProvider = new(tenantId, clientId, clientSecret);

    AccessToken myToken = authProvider.GetToken(
            new TokenRequestContext(["https://graph.microsoft.com/.default"]));

    GraphServiceClient graphClient = new(authProvider);

    return graphClient;
}
//gavdcodeend 016

//gavdcodebegin 017
static GraphServiceClient CsEntraGraphCsSdk_LoginWithManagedIdentitySystem_ForAzFuncts()
{
    DefaultAzureCredential authProvider = new();

    AccessToken myToken = authProvider.GetToken(
            new TokenRequestContext(["https://graph.microsoft.com/.default"]));

    GraphServiceClient graphClient = new(authProvider);

    return graphClient;
}
//gavdcodeend 017

//gavdcodebegin 018
static GraphServiceClient CsEntraGraphCsSdk_LoginWithManagedIdentityUser_ForAzFuncts(
                                                        string clientId)
{
    DefaultAzureCredential authProvider = new(
        new DefaultAzureCredentialOptions { 
            ManagedIdentityClientId = clientId });

    AccessToken myToken = authProvider.GetToken(
            new TokenRequestContext(["https://graph.microsoft.com/.default"]));

    GraphServiceClient graphClient = new(authProvider);

    return graphClient;
}
//gavdcodeend 018

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsEntraGraphCsSdk_GetAllAppRegistrations()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    ApplicationCollectionResponse allApps = myGraphClient.Applications.GetAsync().Result;

    foreach (Application oneApp in allApps.Value)
    {
        Console.WriteLine(oneApp.DisplayName + " - " + oneApp.Id);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsEntraGraphCsSdk_GetOneAppRegistrationByObjectId()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    // Use the Object ID, not the Client ID
    string myAppObjectId = "824741c8-xxxx-xxxx-xxxx-a2d0181cd1c4";

    Application oneApps = myGraphClient.Applications[myAppObjectId].GetAsync().Result;

    Console.WriteLine(oneApps.DisplayName);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsEntraGraphCsSdk_GetOneAppRegistrationByClientId()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    // Use the Client ID, not the Object ID
    string myAppClientId = "5a84f9ed-xxxx-xxxx-xxxx-42efb58acd2a";

    Application oneApp = myGraphClient.ApplicationsWithAppId(myAppClientId)
                                       .GetAsync().Result;

    Console.WriteLine(oneApp.DisplayName);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsEntraGraphCsSdk_GetOneAppRegistrationByObjectIdByProperties()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    // Use the Object ID, not the Client ID
    string myAppObjectId = "824741c8-xxxx-xxxx-xxxx-a2d0181cd1c4";

    Application oneApp = myGraphClient
                .Applications[myAppObjectId].GetAsync((requestConfiguration) =>
    {
        requestConfiguration
                .QueryParameters.Select = ["displayName", "appId", "id"];
    }).Result;

    Console.WriteLine(oneApp.DisplayName);
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsEntraGraphCsSdk_CreateAppRegistrationGraphApi()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    Application myBody = new()
    {
        DisplayName = "Test_MyAppRegFromGraphSdk",
    };
    Application myApp = myGraphClient.Applications.PostAsync(myBody).Result;

    Console.WriteLine(myApp.AppId + " - " + myApp.DisplayName);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsEntraGraphCsSdk_AddOwnerToAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myAppObjectId = "a4b7596b-xxxx-xxxx-xxxx-0c3ff8ff551d"; // Object ID
    string myUserEmail = "user@domain.onmicrosoft.com";

    // Find the User ID by Email
    User myUser = myGraphClient.Users[myUserEmail].GetAsync().Result;
    string myUserId = myUser.Id;
    Console.WriteLine("User ID - " + myUserId);

    // Create a Service Principal for the Application
    ServicePrincipal requestBody = new()
    {
        AppId = myAppClientId,
    };

    ServicePrincipal myResponse = myGraphClient.ServicePrincipals
                                               .PostAsync(requestBody).Result;
    Console.WriteLine("Service Principal ID - " + myResponse.Id);

    // Add the User as an Owner of the App Registration
    ReferenceCreate myBody = new()
    {
        OdataId = "https://graph.microsoft.com/v1.0/directoryObjects/" + myUserId,
    };

    myGraphClient.Applications[myAppObjectId].Owners.Ref.PostAsync(myBody);

    Console.WriteLine("User set as owner");
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsEntraGraphCsSdk_AddDelegatedClaimsToAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myClaimName = "User.ReadWrite.All";

    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    // Get the service principal for Microsoft Graph
    ServicePrincipalCollectionResponse graphServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "displayName eq 'Microsoft Graph'";
                    }).Result;

    // Get the oAuth2PermissionScope for Delegated
    var claimScope = graphServicePrincipal.Value[0].Oauth2PermissionScopes
                            .FirstOrDefault(scope => scope.Value == myClaimName);

    // Grant the Delegated permission
    OAuth2PermissionGrant myOAuth2PermissionGrant = new()
    {
        ClientId = clientServicePrincipal.Value[0].Id,
        ConsentType = "AllPrincipals",
        PrincipalId = null,
        ResourceId = graphServicePrincipal.Value[0].Id,
        Scope = claimScope.Value
    };
    myGraphClient.Oauth2PermissionGrants.PostAsync(myOAuth2PermissionGrant).Wait();

    Console.WriteLine("Delegated permission granted");
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsEntraGraphCsSdk_DeleteDelegatedClaimsFromAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myClaimName = "User.ReadWrite.All";

    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    // Get the scope for the Delegated claim
    OAuth2PermissionGrantCollectionResponse myPermissionGrants =
                            myGraphClient.Oauth2PermissionGrants.GetAsync().Result;
    OAuth2PermissionGrant myClaim = myPermissionGrants.Value
                .Where(claim => claim.ClientId == clientServicePrincipal.Value[0].Id &&
                                claim.Scope == myClaimName).FirstOrDefault();

    // Delete the Delegated permission
    myGraphClient.Oauth2PermissionGrants[myClaim.Id].DeleteAsync().Wait();

    Console.WriteLine("Delegated permission deleted");
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsEntraGraphCsSdk_AddApplicationClaimsToAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myClaimName = "AuditLog.Read.All";

    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    // Get the service principal for Microsoft Graph
    ServicePrincipalCollectionResponse graphServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "displayName eq 'Microsoft Graph'";
                    }).Result;

    // Get the Role
    var claimRole = graphServicePrincipal.Value[0].AppRoles
                            .FirstOrDefault(scope => scope.Value == myClaimName);

    // Grant the Application permission
    AppRoleAssignment myAppRoleAssignment = new()
    {
        PrincipalId = Guid.Parse(clientServicePrincipal.Value[0].Id),
        ResourceId = Guid.Parse(graphServicePrincipal.Value[0].Id),
        AppRoleId = claimRole.Id
    };

    // Grant the Application permission
    myGraphClient.ServicePrincipals[clientServicePrincipal.Value[0].Id]
                            .AppRoleAssignedTo.PostAsync(myAppRoleAssignment).Wait();

    Console.WriteLine("Application permission granted");
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsEntraGraphCsSdk_DeleteApplicationClaimsFromAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myClaimName = "AuditLog.Read.All";

    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    // Get the service principal for Microsoft Graph
    ServicePrincipalCollectionResponse graphServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "displayName eq 'Microsoft Graph'";
                    }).Result;

    // Get the oAuth2PermissionScope for Application
    AppRole myGraphRole = graphServicePrincipal.Value[0].AppRoles
                            .FirstOrDefault(scope => scope.Value == myClaimName);

    // Get the Application Role Assignment
    AppRoleAssignmentCollectionResponse myAllRoles = myGraphClient
                        .ServicePrincipals[clientServicePrincipal.Value[0].Id]
                        .AppRoleAssignments.GetAsync().Result;
    AppRoleAssignment myRole = myAllRoles.Value
                        .FirstOrDefault(role => role.AppRoleId == myGraphRole.Id);

    // Delete the Application permission
    myGraphClient.ServicePrincipals[clientServicePrincipal.Value[0].Id]
                        .AppRoleAssignments[myRole.Id].DeleteAsync().Wait();

    Console.WriteLine("Application permission deleted");
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsEntraGraphCsSdk_AddSecretToAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppObjectId = "a4b7596b-xxxx-xxxx-xxxx-0c3ff8ff551d"; // Object ID

    // The values for the Secret
    string mySecretName = "My AppReg Secret";
    int mySecretDurationInMonths = 26;

    // Add the Secret to the App Registration
    AddPasswordPostRequestBody myBody = new()
    {
        PasswordCredential = new PasswordCredential
        {
            DisplayName = mySecretName,
            EndDateTime = DateTime.Now.AddMonths(mySecretDurationInMonths)
        }
    };

    PasswordCredential myResponse = myGraphClient.Applications[myAppObjectId]
                                                 .AddPassword.PostAsync(myBody).Result;

    Console.WriteLine("Secret - " + myResponse.SecretText);
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsEntraGraphCsSdk_DeleteSecretFromAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppObjectId = "a4b7596b-xxxx-xxxx-xxxx-0c3ff8ff551d"; // Object ID

    // The values for the Secret
    string mySecretName = "My AppReg Secret";

    // Get the application details
    Application myAppRegistration = myGraphClient.Applications[myAppObjectId].GetAsync().Result;

    // Find the secret to remove
    PasswordCredential mySecret = myAppRegistration.PasswordCredentials
                    .Where(secret => secret.DisplayName == mySecretName).FirstOrDefault();

    // Remove the secret
    RemovePasswordPostRequestBody myBody = new()
    {
        KeyId = mySecret.KeyId
    };

    myGraphClient.Applications[myAppObjectId].RemovePassword.PostAsync(myBody).Wait();

    Console.WriteLine("Secret deleted");
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsEntraGraphCsSdk_AddCertificateToAppRegistration()
{
    // Requires Application.Read.All, AppRoleAssignment.ReadWrite.All,
    // Application.ReadWrite.OwnedBy and Directory.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];

    string myCertPathPublic = @"C:\Temporary\MyCertificate.cer";
    string myCertPathPrivate = @"C:\Temporary\MyCertificate.pfx";
    string myCertPrivatePwd = "MyPassword";
    string myCertName = "CN=MyGraphApiCert";
    string myCertFriendlyName = "My Graph Api Cert";
    int myCertDurationInMonths = 23;  // Max duration: 24 months

    using RSA rsa = RSA.Create();
    CertificateRequest myCertificateRequest = new(
        myCertName,
        rsa,
        HashAlgorithmName.SHA256,
        RSASignaturePadding.Pkcs1);

    myCertificateRequest.CertificateExtensions.Add(
        new X509BasicConstraintsExtension(false, false, 0, false));

    X509Certificate2 myCertificate = myCertificateRequest.CreateSelfSigned(
        DateTimeOffset.UtcNow.AddDays(-1),
        DateTimeOffset.UtcNow.AddMonths(myCertDurationInMonths));

    myCertificate.FriendlyName = myCertFriendlyName;

    // Save the certificate to the Windows certificate winStore (Cert:\CurrentUser\My)
    using (X509Store winStore = new(StoreName.My, StoreLocation.CurrentUser))
    {
        winStore.Open(OpenFlags.ReadWrite);
        winStore.Add(myCertificate);
        winStore.Close();
    }

    // Export the Certificate public key to a file
    File.WriteAllBytes(myCertPathPublic, myCertificate.Export(X509ContentType.Cert));
    string myCertB64 = Convert.ToBase64String(File.ReadAllBytes(myCertPathPublic));

    // Export the Certificate private key to a file
    byte[] privateKeyBytes = myCertificate.Export(
                            X509ContentType.Pfx, myCertPrivatePwd);
    File.WriteAllBytes(myCertPathPrivate, privateKeyBytes);

    // Get the Certificate's Thumbprint Base64 Strings
    string myCertThumbprint = myCertificate.Thumbprint;
    string myCertThumbprintB64 = Convert.ToBase64String(
                            System.Text.Encoding.UTF8.GetBytes(myCertThumbprint));

    // Find the "Proof" property for the certificate
    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    X509Certificate2 privCertificate = new(myCertPathPrivate, myCertPrivatePwd);
    X509SecurityKey securityKey = new(privCertificate);

    JwtSecurityTokenHandler tokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Issuer = clientServicePrincipal.Value[0].Id,
        Audience = $"https://graph.microsoft.com/{myTenantId}",
        Claims = new Dictionary<string, object>
            {
                { "keyid", clientServicePrincipal.Value[0].Id }
            },
        Expires = DateTime.UtcNow.AddMinutes(5),
        SigningCredentials = new SigningCredentials(
                                securityKey, SecurityAlgorithms.RsaSha256)
    };

    SecurityToken myToken = tokenHandler.CreateToken(tokenDescriptor);
    string myProof = tokenHandler.WriteToken(myToken);

    // Add the Certificate to the App Registration
    Microsoft.Graph.ServicePrincipals.Item.AddKey.AddKeyPostRequestBody myBody = new()
    {
        KeyCredential = new KeyCredential
        {
            DisplayName = myCertFriendlyName,
            Type = "AsymmetricX509Cert",
            Usage = "Verify",
            EndDateTime = DateTime.Now.AddMonths(myCertDurationInMonths),
            Key = Convert.FromBase64String(myCertB64)
        },
        PasswordCredential = null,
        Proof = myProof
    };

    myGraphClient.ServicePrincipals[clientServicePrincipal.Value[0].Id]
                                                 .AddKey.PostAsync(myBody).Wait();

    Console.WriteLine("Thumbprint: " + myCertThumbprint);
    Console.WriteLine("Thumbprint Base64: " + myCertThumbprintB64);
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsEntraGraphCsSdk_DeleteCertificateFromAppRegistrationAndComputer()
{
    // Requires Application.Read.All, AppRoleAssignment.ReadWrite.All,
    // Application.ReadWrite.OwnedBy and Directory.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    string myAppObjectId = "a4b7596b-xxxx-xxxx-xxxx-0c3ff8ff551d"; // Object ID
    string myAppClientId = "d86afffc-xxxx-xxxx-xxxx-6ddd9a347033"; // Client ID
    string myCertPathPrivate = @"C:\Temporary\MyCertificate.pfx";
    string myCertPrivatePwd = "MyPassword";
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];

    // The values for the Certificate Thumbprint
    X509Certificate2 privCertificate = new(myCertPathPrivate, myCertPrivatePwd);
    X509SecurityKey securityKey = new(privCertificate);
    string myCertThumbp = privCertificate.Thumbprint;

    // Get the application details
    Application myAppRegistration = myGraphClient
                    .Applications[myAppObjectId].GetAsync().Result;

    // Find the certificate to remove
    KeyCredential myCert = myAppRegistration.KeyCredentials
                    .Where(crt => crt.CustomKeyIdentifier != null &&
                           crt.CustomKeyIdentifier.SequenceEqual(
                               Convert.FromBase64String(myCertThumbp))).FirstOrDefault();

    // Find the "Proof" property for the certificate
    // Get the client service principal
    ServicePrincipalCollectionResponse clientServicePrincipal =
                    myGraphClient.ServicePrincipals.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter =
                                    "appId eq '" + myAppClientId + "'";
                    }).Result;

    JwtSecurityTokenHandler tokenHandler = new();
    SecurityTokenDescriptor tokenDescriptor = new()
    {
        Issuer = clientServicePrincipal.Value[0].Id,
        Audience = $"https://graph.microsoft.com/{myTenantId}",
        Claims = new Dictionary<string, object>
            {
                { "keyid", clientServicePrincipal.Value[0].Id }
            },
        Expires = DateTime.UtcNow.AddMinutes(5),
        SigningCredentials = new SigningCredentials(securityKey,
                                                    SecurityAlgorithms.RsaSha256)
    };

    SecurityToken myToken = tokenHandler.CreateToken(tokenDescriptor);
    string myProof = tokenHandler.WriteToken(myToken);

    // Remove the certificate
    Microsoft.Graph.ServicePrincipals.Item.RemoveKey
        .RemoveKeyPostRequestBody myBody = new()
        {
            KeyId = myCert.KeyId,
            Proof = myProof
        };

    myGraphClient.ServicePrincipals[clientServicePrincipal.Value[0].Id]
                                        .RemoveKey.PostAsync(myBody).Wait();

    Console.WriteLine("Certificate deleted");

    // Delete Certificate from the Windows Certificate Store
    // Access the Local Machine's My store
    X509Store winStore = new("My", StoreLocation.CurrentUser);
    winStore.Open(OpenFlags.ReadWrite);

    // Find the certificate by thumbprint
    X509Certificate2 myCertLocal = winStore.Certificates
                        .Find(X509FindType.FindByThumbprint, myCertThumbp, false)
                        .OfType<X509Certificate2>()
                        .FirstOrDefault();

    winStore.Remove(myCertLocal);

    Console.WriteLine("Certificate removed from the Windows Certificate Store");

    // Delete the Certificate files if necessary
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsEntraGraphCsSdk_DeleteAppRegistration()
{
    // Requires Application.Read.All and AppRoleAssignment.ReadWrite.All

    GraphServiceClient myGraphClient = CsEntraGraphCsSdk_LoginWithSecret();

    // Use the Object ID, not the Client ID
    string myAppObjectId = "a4b7596b-xxxx-xxxx-xxxx-0c3ff8ff551d";

    myGraphClient.Applications[myAppObjectId].DeleteAsync().Wait();

    Console.WriteLine("App Registration deleted");
}
//gavdcodeend 015

// Routine that can be used in Azure Functions with Managed Identities
//gavdcodebegin 019
static List<string> GetSharePointDocs_ForAzureFunctions(string siteId, string clientId,
                                    string tenantId, string clientSecret)
{
    //GraphServiceClient myGraphClient =
    //  CsEntraGraphCsSdk_LoginWithSecret_ForAzureFunctions(clientId, tenantId, clientSecret);
    //GraphServiceClient myGraphClient = 
    //    CsEntraGraphCsSdk_LoginWithManagedIdentitySystem_ForAzureFunctions();
    GraphServiceClient myGraphClient = 
        CsEntraGraphCsSdk_LoginWithManagedIdentityUser_ForAzureFunctions(clientId);

    //ListCollectionResponse lists = myGraphClient.Sites[siteId].Lists.GetAsync().Result;

    ListItemCollectionResponse allDocs = myGraphClient
        .Sites[siteId]
        .Lists["Documents"].Items
        .GetAsync().Result;
    List<string> myDocs = [];
    foreach (ListItem oneDoc in allDocs.Value)
    {
        myDocs.Add(oneDoc.WebUrl);
    }

    return myDocs;
}
//gavdcodeend 019

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 019 ***

//CsEntraGraphCsSdk_GetAllAppRegistrations();
//CsEntraGraphCsSdk_GetOneAppRegistrationByObjectId();
//CsEntraGraphCsSdk_GetOneAppRegistrationByClientId();
//CsEntraGraphCsSdk_GetOneAppRegistrationByObjectIdByProperties();
//CsEntraGraphCsSdk_CreateAppRegistrationGraphApi();
//CsEntraGraphCsSdk_AddOwnerToAppRegistration();
//CsEntraGraphCsSdk_AddDelegatedClaimsToAppRegistration();
//CsEntraGraphCsSdk_DeleteDelegatedClaimsFromAppRegistration();
//CsEntraGraphCsSdk_AddApplicationClaimsToAppRegistration();
//CsEntraGraphCsSdk_DeleteApplicationClaimsFromAppRegistration();
//CsEntraGraphCsSdk_AddSecretToAppRegistration();
//CsEntraGraphCsSdk_DeleteSecretFromAppRegistration();
//CsEntraGraphCsSdk_AddCertificateToAppRegistration();
//CsEntraGraphCsSdk_DeleteCertificateFromAppRegistrationAndComputer();
//CsEntraGraphCsSdk_DeleteAppRegistration();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
#pragma warning restore CA1416 // Validate platform compatibility
