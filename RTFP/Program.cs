using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

const string myTenantId = "ade56059-89c0-4594-90c3-e4772a8168ca";
const string myClientId = "3d16c7bc-a14e-454e-81cd-da571cf2a8e3";
const string myClientSecret = "IIX8Q~M5q-FCotlzO7GA4bLAQBPtF9dHwr.CcbsN";

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
static void CsSpGraphSdk_GetAllSpEmbeddedContainers()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    FileStorageContainerCollectionResponse allContainers =
        myGraphClient.Storage.FileStorage.Containers.GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Filter =
                    "containerTypeId eq " + myContainerTypeId;
    }).Result;

    foreach (FileStorageContainer oneContainer in allContainers.Value)
    {
        Console.WriteLine(oneContainer.DisplayName + " - " + oneContainer.Id);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpGraphSdk_GetOneSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9...6RqiYjT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    FileStorageContainer myContainer = myGraphClient.Storage.FileStorage
                                        .Containers[myContainerId].GetAsync().Result;

    Console.WriteLine(myContainer.Status);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpGraphSdk_GetDriveSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9...6RqiYjT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Drive myContainer = myGraphClient.Storage.FileStorage
                                    .Containers[myContainerId].Drive.GetAsync().Result;

    Console.WriteLine(myContainer.LastModifiedDateTime);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpGraphSdk_ActivateSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9...6RqiYjT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient.Storage.FileStorage.Containers[myContainerId]
                                        .Activate.PostAsync().Wait();
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpGraphSdk_CreateSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    FileStorageContainer myBody = new()
    {
        DisplayName = "Another Test Storage Container",
        Description = "It is Another Test Storage Container",
        ContainerTypeId = Guid.Parse(myContainerTypeId),
    };

    FileStorageContainer myContainer = myGraphClient.Storage.FileStorage
                                            .Containers.PostAsync(myBody).Result;

    Console.WriteLine(myContainer.Id);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpGraphSdk_UpdateSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    FileStorageContainer myBody = new()
    {
        DisplayName = "Another Test Storage Container Updated",
        Description = "It is Another Test Storage Container Updated"
    };

    FileStorageContainer myContainer = myGraphClient.Storage.FileStorage
                                .Containers[myContainerId].PatchAsync(myBody).Result;

    Console.WriteLine(myContainer.Id);
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpGraphSdk_DeleteSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient.Storage.FileStorage.Containers[myContainerId].DeleteAsync().Wait();
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpGraphSdk_DeletePermanentlySpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient.Storage.FileStorage.Containers[myContainerId]
                                                .PermanentDelete.PostAsync().Wait();
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpGraphSdk_GetAllPermissionsSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    PermissionCollectionResponse allPermissions = myGraphClient.Storage.FileStorage
                            .Containers[myContainerId].Permissions.GetAsync().Result;

    foreach (Permission onePermission in allPermissions.Value)
    {
        Console.WriteLine(onePermission.GrantedToV2.User.DisplayName + 
                            " - " + onePermission.Id);
    }
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpGraphSdk_AddPermissionSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Permission requestBody = new()
    {
        Roles = new List<string>
        {
            "reader",
        },
        GrantedToV2 = new SharePointIdentitySet
        {
            User = new Identity
            {
                AdditionalData = new Dictionary<string, object>
                {
                    {
                        "userPrincipalName" , "user@domain.onmicrosoft.com"
                    },
                },
            },
        }
    };

    Permission myPermission = myGraphClient.Storage.FileStorage
                .Containers[myContainerId].Permissions.PostAsync(requestBody).Result;

    Console.WriteLine(myPermission.Id);
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpGraphSdk_UpdatePermissionSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myPermissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZW...ubWljcm9zb2Z0LmNvbQ";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Permission requestBody = new()
    {
        Roles = new List<string>
        {
            "owner",
        }
    };

    Permission myPermission = myGraphClient.Storage.FileStorage
                        .Containers[myContainerId].Permissions[myPermissionId]
                        .PatchAsync(requestBody).Result;

    Console.WriteLine(myPermission.GrantedToV2.User.DisplayName);
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpGraphSdk_DeletePermissionSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myPermissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZW...ubWljcm9zb2Z0LmNvbQ";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient.Storage.FileStorage
                        .Containers[myContainerId].Permissions[myPermissionId]
                        .DeleteAsync().Wait();
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpGraphSdk_UploadOneFileToLibrary()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myFilePath = @"C:\Temporary\TestWordFile.docx";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    using var fileStream = new FileStream(myFilePath, FileMode.Open); 
    DriveItem uploadedFile = myGraphClient
                        .Drives[myContainerId]
                        .Root
                        .ItemWithPath(Path.GetFileName(myFilePath))
                        .Content
                        .PutAsync(fileStream).Result; 
    
    Console.WriteLine("File ID: " + uploadedFile.Id);
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpGraphSdk_GetAllFilesInFolderLibrary()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    DriveItem myRoot = myGraphClient
                        .Drives[myContainerId]
                        .Root.GetAsync().Result;

    DriveItemCollectionResponse myDocs = myGraphClient
                        .Drives[myContainerId]
                        .Items[myRoot.Id]
                        .Children
                        .GetAsync().Result;

    foreach (DriveItem oneDoc in myDocs.Value)
    {
        Console.WriteLine(oneDoc.Name + " - " + oneDoc.Id);
    }
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpGraphSdk_DownloadFileFromSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myFileId = "01OOC26X34R32PFMWJCNCZZXADTFWDFHUF";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    Stream myFile = myGraphClient
                        .Drives[myContainerId]
                        .Items[myFileId]
                        .Content
                        .GetAsync().Result;

    string myFilePath = @"C:\Temporary\TestWordFile(Download).docx"; 
    using FileStream fileStream = new(myFilePath, FileMode.Create); 
    myFile.CopyTo(fileStream);
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpGraphSdk_GetAllMetadataItemSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myFileId = "01OOC26X34R32PFMWJCNCZZXADTFWDFHUF";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    DriveItem myFile = myGraphClient
                        .Drives[myContainerId]
                        .Items[myFileId]
                        .GetAsync().Result;

    Console.WriteLine(myFile.Name); 
    Console.WriteLine(myFile.Id); 
    Console.WriteLine(myFile.Size + " bytes"); 
    Console.WriteLine(myFile.CreatedDateTime); 
    Console.WriteLine(myFile.LastModifiedDateTime);
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsSpGraphSdk_DeleteFileFromSpEmbeddedContainer()
{
    // Requires Delegated rights: FileStorageContainer.Selected

    string myContainerId = "b!KC8lA9xDLkC4u111xWn5qM_2HzkZmN9Gonxt4GL...jT8Qtn88";
    string myFileId = "01OOC26X34R32PFMWJCNCZZXADTFWDFHUF";

    GraphServiceClient myGraphClient =
            CsSpGraphSdk_LoginWithSecret(myTenantId, myClientId, myClientSecret);

    myGraphClient
            .Drives[myContainerId]
            .Items[myFileId]
            .DeleteAsync().Wait();
}
//gavdcodeend 017

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 017 ***

//CsSpGraphSdk_GetAllSpEmbeddedContainers();
//CsSpGraphSdk_GetOneSpEmbeddedContainer();
//CsSpGraphSdk_GetDriveSpEmbeddedContainer();
//CsSpGraphSdk_ActivateSpEmbeddedContainer();
//CsSpGraphSdk_CreateSpEmbeddedContainer();
//CsSpGraphSdk_UpdateSpEmbeddedContainer();
//CsSpGraphSdk_DeleteSpEmbeddedContainer();
//CsSpGraphSdk_DeletePermanentlySpEmbeddedContainer();
//CsSpGraphSdk_GetAllPermissionsSpEmbeddedContainer();
//CsSpGraphSdk_AddPermissionSpEmbeddedContainer();
//CsSpGraphSdk_UpdatePermissionSpEmbeddedContainer();
//CsSpGraphSdk_DeletePermissionSpEmbeddedContainer();
//CsSpGraphSdk_UploadOneFileToLibrary();
//CsSpGraphSdk_GetAllFilesInFolderLibrary();
//CsSpGraphSdk_DownloadFileFromSpEmbeddedContainer();
//CsSpGraphSdk_GetAllMetadataItemSpEmbeddedContainer();
//CsSpGraphSdk_DeleteFileFromSpEmbeddedContainer();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
