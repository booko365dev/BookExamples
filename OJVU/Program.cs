using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.Checkin;
using Microsoft.Graph.Drives.Item.Items.Item.Copy;
using Microsoft.Graph.Drives.Item.Items.Item.Invite;
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
static void CsSpGraphSdk_GetAllItemsInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ListItemCollectionResponse myItems = myGraphClient
                                    .Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                    .Lists["c6d81938-b786-4af4-b2dd-a2132787f1d9"]
                                    .Items
                                    .GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Expand = ["fields($select=Title)"];
        }).Result;

    foreach (ListItem oneItem in myItems.Value)
    {
        Console.WriteLine(oneItem.Fields.AdditionalData["Title"] + " - " + oneItem.Id);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpGraphSdk_GetOneListItemInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    ListItem myItem = myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                   .Lists["c6d81938-b786-4af4-b2dd-a2132787f1d9"]
                                   .Items["11"]
                                   .GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.Expand = ["fields($select=Title)"];
        }).Result;

    Console.WriteLine(myItem.Fields.AdditionalData["Title"] + " - " + myItem.Id);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpGraphSdk_CreateOneListItemInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    var myBody = new ListItem
    {
        Fields = new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object>
            {
                {
                    "Title" , "PsSpGraphSDK_Item"
                }
            }
        }
    };

    ListItem myItem = myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                   .Lists["c6d81938-b786-4af4-b2dd-a2132787f1d9"]
                                   .Items
                                   .PostAsync(myBody).Result;

    Console.WriteLine(myItem.Id);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpGraphSdk_UpdateOneListItemInList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    var myBody = new ListItem
    {
        Fields = new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object>
            {
                {
                    "Title" , "Update_PsSpGraphSDK_Item"
                }
            }
        }
    };

    ListItem myItem = myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                   .Lists["c6d81938-b786-4af4-b2dd-a2132787f1d9"]
                                   .Items["13"]
                                   .PatchAsync(myBody).Result;
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpGraphSdk_DeleteOneListItemFromList()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                 .Lists["c6d81938-b786-4af4-b2dd-a2132787f1d9"]
                 .Items["13"]
                 .DeleteAsync().Wait();
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpGraphSdk_GetDriveOfLibraryByLibraryId()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Drive myListDrive = myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                     .Lists["b331af17-edc3-4058-909e-e6fa74abe946"]
                                     .Drive
                                     .GetAsync().Result;

    Console.WriteLine(myListDrive.DriveType + " - " + myListDrive.Id);
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpGraphSdk_GetAllFilesInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Drive myDrive = myGraphClient.Sites["91ee115a-8a5b-49ad-9627-99dae04394ab"]
                                 .Lists["b331af17-edc3-4058-909e-e6fa74abe946"]
                                 .Drive
                                 .GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Expand = ["root"];
                    }).Result;

    string myLibraryRootId = myDrive.Root.Id;
    Console.WriteLine("Root ID of Library - " + myLibraryRootId);
    DriveItemCollectionResponse myDocs = myGraphClient.Drives[myDrive.Id]
                                                      .Items[myLibraryRootId]
                                                      .Children
                                                      .GetAsync().Result;

    foreach (DriveItem oneDoc in myDocs.Value)
    {
        Console.WriteLine(oneDoc.Name + " - " + oneDoc.Id);
    }
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpGraphSdk_GetAllFilesInFolderLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    DriveItemCollectionResponse myDocs = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RFCB4MZLKVESFH2OPTWWU7NJHOX"]
        .Children
        .GetAsync().Result;

    foreach (DriveItem oneDoc in myDocs.Value)
    {
        Console.WriteLine(oneDoc.Name + " - " + oneDoc.Id);
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpGraphSdk_GetOneFileMetadata()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    DriveItem myDoc = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RAGP74HVDQYU5DK5TRD3W7XVEOU"]
        .GetAsync().Result;

    Console.WriteLine(myDoc.Name);
    Console.WriteLine(myDoc.Id);
    Console.WriteLine(myDoc.Size + " bytes");
    Console.WriteLine(myDoc.CreatedDateTime);
    Console.WriteLine(myDoc.LastModifiedDateTime);
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpGraphSdk_UploadOneFileToLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    string myFilePath = @"C:\Temporary\TestDocument.docx";

    using var fileStream = new FileStream(myFilePath, FileMode.Open);
    DriveItem uploadedFile = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Root
        .ItemWithPath(Path.GetFileName(myFilePath))
        .Content
        .PutAsync(fileStream).Result;

    Console.WriteLine("File ID: " + uploadedFile.Id);
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpGraphSdk_DownloadOneFileFromLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Stream myDoc = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Content
        .GetAsync().Result;

    string myFilePath = @"C:\Temporary\TestDocument(Download).docx";
    using FileStream fileStream = new FileStream(myFilePath, FileMode.Create);
    myDoc.CopyTo(fileStream);
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpGraphSdk_UpdateOneFileMetadata()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    DriveItem myBody = new()
    {
        Name = "TestDocument(Updated).docx",
    };

    DriveItem myDoc = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .PatchAsync(myBody).Result;
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpGraphSdk_DeleteOneFile()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .DeleteAsync().Wait();
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpGraphSdk_CopyFile()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    string myNewDriveId = "b!WhHukVuKrUmWJ5na4EInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm";
    Drive myNewDrive = myGraphClient.Drives[myNewDriveId]
                            .GetAsync((requestConfiguration) =>
                            {
                                requestConfiguration.QueryParameters.Expand = ["root"];
                            }).Result;

    string myNewLibraryRootId = myNewDrive.Root.Id;

    CopyPostRequestBody myBody = new()
    {
        ParentReference = new ItemReference
        {
            DriveId = myNewDriveId,
            Id = myNewLibraryRootId
        },
        Name = "TestDocument(Copy).docx",
    };

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Copy.PostAsync(myBody);
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpGraphSdk_MoveFile()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    string myNewDriveId = "b!WhHukVuKrUmWJ5na4EAYoJInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm";
    Drive myNewDrive = myGraphClient.Drives[myNewDriveId]
                            .GetAsync((requestConfiguration) =>
                            {
                                requestConfiguration.QueryParameters.Expand = ["root"];
                            }).Result;

    string myNewLibraryRootId = myNewDrive.Root.Id;

    DriveItem myBody = new()
    {
        ParentReference = new ItemReference
        {
            DriveId = myNewDriveId,
            Id = myNewLibraryRootId
        },
        Name = "TestDocument(Copy).docx",
    };

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .PatchAsync(myBody);
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpGraphSdk_CreateFolderInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    string myDriveId = "b!WhHukVuKrUmWJ5na4EOUqInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG";
    Drive myNewDrive = myGraphClient.Drives[myDriveId]
                            .GetAsync((requestConfiguration) =>
                            {
                                requestConfiguration.QueryParameters.Expand = ["root"];
                            }).Result;

    string myLibraryRootId = myNewDrive.Root.Id;

    DriveItem myBody = new()
    {
        Name = "NewFolderGraphSDK",
        Folder = new Folder
        {
        },
        AdditionalData = new Dictionary<string, object>
        {
            {
                "@microsoft.graph.conflictBehavior" , "rename"
            }
        }
    };

    myGraphClient
        .Drives[myDriveId]
        .Items[myLibraryRootId]
        .Children
        .PostAsync(myBody);
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsSpGraphSdk_CheckOutFileInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Checkout
        .PostAsync().Wait();
}
//gavdcodeend 017

//gavdcodebegin 018
static void CsSpGraphSdk_CheckInFileInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    CheckinPostRequestBody myBody = new()
    {
        Comment = "Check in from the Graph SDK"
    };

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Checkin
        .PostAsync(myBody).Wait();
}
//gavdcodeend 018

//gavdcodebegin 019
static void CsSpGraphSdk_GetPermissionsFileInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    PermissionCollectionResponse myPermissions = myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Permissions
        .GetAsync().Result;

    foreach (Permission onePermission in myPermissions.Value)
    {
        string rolesString = string.Join(", ", onePermission.Roles);
        Console.WriteLine(onePermission.Id + " - " + rolesString);
    }
}
//gavdcodeend 019

//gavdcodebegin 020
static void CsSpGraphSdk_CreatePermissionFileInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    InvitePostRequestBody myBody = new()
    {
        Recipients = new List<DriveRecipient>
        {
            new DriveRecipient
            {
                Email = "user@domain.onmicrosoft.com",
            },
        },
            Message = "This is a file with permissions",
            RequireSignIn = true,
            SendInvitation = true,
            Roles = new List<string>
        {
            "write",
        },
        Password = "password123",
        ExpirationDateTime = "2024-12-31T23:59:00.000Z",
    };

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Invite
        .PostAsInvitePostResponseAsync(myBody);
}
//gavdcodeend 020

//gavdcodebegin 021
static void CsSpGraphSdk_DeletePermissionFileInLibrary()
{
    // Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string myClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
        .Drives["b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"]
        .Items["01IAJF3RGMALWCFKWMYBBJLMLNOPX4WIXE"]
        .Permissions["aTowIy5mfG1lbWJlcnNoaXBpdGFjYWRldi5vbm1pY3Jvc29mdC5jb20"]
        .DeleteAsync();
}
//gavdcodeend 021


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 021 ***

//CsSpGraphSdk_GetAllItemsInList();
//CsSpGraphSdk_GetOneListItemInList();
//CsSpGraphSdk_CreateOneListItemInList();
//CsSpGraphSdk_UpdateOneListItemInList();
//CsSpGraphSdk_DeleteOneListItemFromList();
//CsSpGraphSdk_GetDriveOfLibraryByLibraryId();
//CsSpGraphSdk_GetAllFilesInLibrary();
//CsSpGraphSdk_GetAllFilesInFolderLibrary();
//CsSpGraphSdk_GetOneFileMetadata();
//CsSpGraphSdk_UploadOneFileToLibrary();
//CsSpGraphSdk_DownloadOneFileFromLibrary();
//CsSpGraphSdk_UpdateOneFileMetadata();
//CsSpGraphSdk_DeleteOneFile();
//CsSpGraphSdk_CopyFile();
//CsSpGraphSdk_MoveFile();
//CsSpGraphSdk_CreateFolderInLibrary();
//CsSpGraphSdk_CheckOutFileInLibrary();
//CsSpGraphSdk_CheckInFileInLibrary();
//CsSpGraphSdk_GetPermissionsFileInLibrary();
//CsSpGraphSdk_CreatePermissionFileInLibrary();
//CsSpGraphSdk_DeletePermissionFileInLibrary();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------




#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
