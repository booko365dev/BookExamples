using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
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
static void CsSpGraphSdk_GetAllListsInSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ListCollectionResponse myList = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists
            .GetAsync().Result;

    foreach (List oneList in myList.Value)
    {
        Console.WriteLine(oneList.WebUrl + " - " + oneList.Id);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpGraphSdk_GetOneListInSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    List myList = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .GetAsync().Result;

    Console.WriteLine(myList.WebUrl);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpGraphSdk_CreateOneListInSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    var myBody = new List
    {
        DisplayName = "List created with Graph SDK",
        Columns =
        [
            new() {
                Name = "MyTextField",
                Text = new TextColumn {  
                }
            },
            new() {
                Name = "MyNumberField",
                Number = new NumberColumn {
                }
            }
        ],
        ListProp = new ListInfo
        {
            Template = "genericList",
        }
    };
    
    List myList = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists
            .PostAsync(myBody).Result;
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpGraphSdk_UpdateOneListInSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    List myBody = new()
    {
        Description = "List updated"
    };

    List myList = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .PatchAsync(myBody).Result;

    Console.WriteLine(myList.Description);
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpGraphSdk_DeleteOneListFromSiteCollection()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .DeleteAsync().Wait();
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpGraphSdk_GetAllColumnsInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ColumnDefinitionCollectionResponse myColumns = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .Columns
            .GetAsync().Result;

    foreach (ColumnDefinition oneColumn in myColumns.Value)
    {
        Console.WriteLine(oneColumn.DisplayName + " - " + oneColumn.Id);
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpGraphSdk_GetOneColumnInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ColumnDefinition myColumn = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .Columns["f1b9501b-a8c8-407b-8395-fc45e01fda05"]
            .GetAsync().Result;

    Console.WriteLine(myColumn.Name);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpGraphSdk_CreateOneColumnInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    var myBody = new ColumnDefinition
    {
        Name = "Column From Graph SDK",
        Description = "Description Column",
        EnforceUniqueValues = false,
        Hidden = false,
        Indexed = false,
        Text = new TextColumn
        {
            AllowMultipleLines = false,
            AppendChangesToExistingText = false,
            LinesForEditing = 0,
            MaxLength = 255
        }
    };

    ColumnDefinition myColumn = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .Columns
            .PostAsync(myBody).Result;

    Console.WriteLine(myColumn.Id);
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpGraphSdk_UpdateOneColumnInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ColumnDefinition myBody = new()
    {
        Description = "Column updated"
    };

    ColumnDefinition myColumn = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .Columns["f1b9501b-a8c8-407b-8395-fc45e01fda05"]
            .PatchAsync(myBody).Result;

    Console.WriteLine(myColumn.Description);
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpGraphSdk_DeleteOneColumnFromList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .Columns["f1b9501b-a8c8-407b-8395-fc45e01fda05"]
            .DeleteAsync().Wait();
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpGraphSdk_GetAllContentTypesInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ContentTypeCollectionResponse myContentTypes = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .ContentTypes
            .GetAsync().Result;

    foreach (ContentType oneContentType in myContentTypes.Value)
    {
        Console.WriteLine(oneContentType.Name + " - " + oneContentType.Id);
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpGraphSdk_GetOneContentTypeInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ContentType myContentType = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .ContentTypes["0x0100C48E769F09462748923DDAAD58A2A72C"]
            .GetAsync().Result;

    Console.WriteLine(myContentType.Name);
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpGraphSdk_CreateOneContentTypeInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    var myBody = new ContentType
    {
        Name = "ContentType From Graph SDK",
        Base = new ContentType
        {
            Id = "0x0101"
        }
    };

    ContentType myContentType = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .ContentTypes
            .PostAsync(myBody).Result;

    Console.WriteLine(myContentType.Id);
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpGraphSdk_UpdateOneContentTypeInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ContentType myBody = new()
    {
        Description = "ContentType updated"
    };

    ContentType myContentType = myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .ContentTypes["0x0100C48E769F09462748923DDAAD58A2A72C"]
            .PatchAsync(myBody).Result;

    Console.WriteLine(myContentType.Description);
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpGraphSdk_DeleteOneContentTypeFromList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
    // Works only for Delegated permissions

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
            .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
            .Lists["ad694da5-3c6d-469f-b529-b1345541ca04"]
            .ContentTypes["0x0100C48E769F09462748923DDAAD58A2A72C"]
            .DeleteAsync().Wait();
}
//gavdcodeend 015


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 015 ***

//CsSpGraphSdk_GetAllListsInSiteCollection();
//CsSpGraphSdk_GetOneListInSiteCollection();
//CsSpGraphSdk_CreateOneListInSiteCollection();
//CsSpGraphSdk_UpdateOneListInSiteCollection();
//CsSpGraphSdk_DeleteOneListFromSiteCollection();
//CsSpGraphSdk_GetAllColumnsInList();
//CsSpGraphSdk_GetOneColumnInList();
//CsSpGraphSdk_CreateOneColumnInList();
//CsSpGraphSdk_UpdateOneColumnInList();
//CsSpGraphSdk_DeleteOneColumnFromList();
//CsSpGraphSdk_GetAllContentTypesInList();
//CsSpGraphSdk_GetOneContentTypeInList();
//CsSpGraphSdk_CreateOneContentTypeInList();
//CsSpGraphSdk_UpdateOneContentTypeInList();
//CsSpGraphSdk_DeleteOneContentTypeFromList();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------




#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
