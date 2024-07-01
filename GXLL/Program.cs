using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Sites.Item.Permissions;
using Microsoft.Graph.Users.Item.FollowedSites.Add;
using Microsoft.Graph.Users.Item.FollowedSites.Remove;
using System.Configuration;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------
static GraphServiceClient CsSpGraphSdk_LoginWithSecret(
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


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpGraphSdk_GetAllSiteCollections()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    SiteCollectionResponse mySites = myGraphClient
            .Sites
            .GetAsync().Result;

    foreach (Site oneSite in mySites.Value)
    {
        Console.WriteLine(oneSite.WebUrl);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpGraphSdk_GetOneSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Site mySite = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .GetAsync().Result;

    Console.WriteLine(mySite.WebUrl);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpGraphSdk_GetFollowedSiteCollections()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    SiteCollectionResponse mySites = myGraphClient
            .Me.FollowedSites
            .GetAsync().Result;

    foreach (Site oneSite in mySites.Value)
    {
        Console.WriteLine(oneSite.WebUrl);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpGraphSdk_FollowSiteCollections()
{
    // Requires Delegated rights: Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    AddPostRequestBody myBody = new()
    {
        Value = [
            new Site
            {
                Id = "domain.sharepoint.com," +
                    "de7cd70c0-6e48-48b4-9380-caab1d1e8433," +
                    "7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
            }
        ],
    };

    AddPostResponse myResult = myGraphClient
            .Users["acc28fcb-5261-47f8-960b-715d2f98a431"]
            .FollowedSites
            .Add
            .PostAsAddPostResponseAsync(myBody)
            .Result;

    foreach (Site oneSite in myResult.Value)
    {
        Console.WriteLine(oneSite.WebUrl);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpGraphSdk_UnFollowSiteCollections()
{
    // Requires Delegated rights: Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    RemovePostRequestBody myBody = new()
    {
        Value = [
            new Site
            {
                Id = "domain.sharepoint.com," +
                    "de7cd70c0-6e48-48b4-9380-caab1d1e8433," +
                    "7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
            }
        ],
    };

    RemovePostResponse myResult = myGraphClient
            .Users["acc28fcb-5261-47f8-960b-715d2f98a431"]
            .FollowedSites
            .Remove
            .PostAsRemovePostResponseAsync(myBody)
            .Result;

    foreach (Site oneSite in myResult.Value)
    {
        Console.WriteLine(oneSite.WebUrl);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpGraphSdk_GetAllPermissionsSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    PermissionCollectionResponse myPermissions = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Permissions
            .GetAsync().Result;

    foreach (Permission onePermission in myPermissions.Value)
    {
        Console.WriteLine(onePermission.Roles + " - " + onePermission.Id);
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpGraphSdk_GetOnePermissionSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Permission myPermission = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Permissions["7a77a5f9-29d8-4fc2-83c7-17c6b8e007af"]
            .GetAsync().Result;

    Console.WriteLine(myPermission.Roles + " - " + myPermission.Id);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpGraphSdk_CreatePermissionSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Permission myBody = new()
    {
        Roles = new List<string>
        {
            "write"
        },
        GrantedTo = new IdentitySet
        {
            Application = new Identity
            {
                Id = "7a77a5f9-29d8-4fc2-83c7-17c6b8e007af",
                DisplayName = "My Permissions"
            }
        }
    };

    Task<Permission> myPermission = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Permissions
            .PostAsync(myBody);
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpGraphSdk_UpdatePermissionSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Permission myBody = new()
    {
        Roles = new List<string>
        {
            "read"
        }
    };

    Task<Permission> myPermission = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Permissions["7a77a5f9-29d8-4fc2-83c7-17c6b8e007af"]
            .PatchAsync(myBody);
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpGraphSdk_DeletePermissionSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Task myPermission = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Permissions["7a77a5f9-29d8-4fc2-83c7-17c6b8e007af"]
            .DeleteAsync();
}
//gavdcodeend 010


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 010 ***

//CsSpGraphSdk_GetAllSiteCollections();
//CsSpGraphSdk_GetOneSiteCollection();
//CsSpGraphSdk_GetFollowedSiteCollections();
//CsSpGraphSdk_FollowSiteCollections();
//CsSpGraphSdk_UnFollowSiteCollections();
//CsSpGraphSdk_GetAllPermissionsSiteCollection();
//CsSpGraphSdk_GetOnePermissionSiteCollection();
//CsSpGraphSdk_CreatePermissionSiteCollection();
//CsSpGraphSdk_UpdatePermissionSiteCollection();
//CsSpGraphSdk_DeletePermissionSiteCollection();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------




#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
