using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext CsSpPnpFramework_LoginWithAccPw()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnpFramework_LoginWithCertificate()
{
    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnpFramework_LoginPnpManagementShell()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnpFramework_LoginWithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query
    ClientContext rtnContext = new
        AuthenticationManager().GetACSAppOnlyContext(
                        ConfigurationManager.AppSettings["SiteCollUrl"],
                        ConfigurationManager.AppSettings["ClientIdWithSecret"],
                        ConfigurationManager.AppSettings["ClientSecret"]);

    return rtnContext;
}

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpPnpFramework_CreatePropertyBag()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

    myList.SetPropertyBagValue("myKey", "myValueString");
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnpFramework_ReadPropertyBag()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

    string myKeyValue = myList.GetPropertyBagValueString("myKey", "");
    Console.WriteLine(myKeyValue);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnpFramework_PropertyBagExists()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

    bool myKeyExists = myList.PropertyBagContainsKey("myKey");
    Console.WriteLine(myKeyExists.ToString());
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnpFramework_PropertyBagIndex()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

    myList.AddIndexedPropertyBagKey("myKey");
    IEnumerable<string> myIndexedPropertyBagKeys =
                                        myList.GetIndexedPropertyBagKeys();

    foreach (string oneKey in myIndexedPropertyBagKeys)
    {
        Console.WriteLine(oneKey);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnpFramework_DeletePropertyBag()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

    myList.RemovePropertyBagValue("myKey");
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnpFramework_DownloadFile()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathRelative =
            "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";
    string pathLocal = @"C:\Temporary";
    string fileName = "TestText.txt";

    spPnpCtx.Web.SaveFileToLocal(pathRelative, pathLocal, fileName);
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnpFramework_DownloadFileAsString()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathRelative =
            "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";
    string myFileAsText = spPnpCtx.Web.GetFileAsString(pathRelative);

    Console.WriteLine(myFileAsText);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnpFramework_FindFiles()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    // Very slow method
    List<Microsoft.SharePoint.Client.File> allFiles = spPnpCtx.Web.FindFiles("*.txt");

    foreach (Microsoft.SharePoint.Client.File oneFile in allFiles)
    {
        Console.WriteLine(oneFile.Name + " - " + oneFile.ServerRelativeUrl);
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpPnpFramework_RequireUpload()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List<Microsoft.SharePoint.Client.File> myFiles =
                                            spPnpCtx.Web.FindFiles("TestText.txt");
    Microsoft.SharePoint.Client.File myFile = myFiles[0];

    // Very slow method
    bool requireUpload = myFile.VerifyIfUploadRequired(@"C:\Temporary\TestText.txt");

    Console.WriteLine("Require Upload " + requireUpload.ToString());
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpPnpFramework_ResetVersion()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathRelative =
            "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";

    spPnpCtx.Web.ResetFileToPreviousVersion(pathRelative,
                                CheckinType.MinorCheckIn, "Done by PnPCore");
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpPnpFramework_CreateFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.CreateFolder("PnPFrameworkFolder");
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpPnpFramework_EnsureFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkEnsureFolder");
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpPnpFramework_CreateSubFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
    Folder mySubFolder = myFolder.EnsureFolder("PnPFrameworkSubFolder");
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpPnpFramework_FolderExistsBool()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    bool fldExists = myList.RootFolder.FolderExists("PnPFrameworkFolder");
    Console.WriteLine("Folder exists - " + fldExists.ToString());
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpPnpFramework_FolderExistsFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder fldExists = myList.RootFolder.ResolveSubFolder("PnPFrameworkFolder");
    Console.WriteLine("Folder exists - " + fldExists.ServerRelativeUrl);
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpPnpFramework_FolderExistsWeb()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathRelative =
            "/sites/[SiteName]/[LibraryName]/[FolderName]";

    bool fldExists = spPnpCtx.Site.RootWeb.DoesFolderExists(pathRelative);

    Console.WriteLine("Folder exists - " + fldExists.ToString());
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsSpPnpFramework_CreateSubFolderFromFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.ResolveSubFolder("PnPFrameworkFolder");
    Folder mySubFolder = myFolder.CreateFolder("PnPFrameworkSubFolder02");
}
//gavdcodeend 017

//gavdcodebegin 018
static void CsSpPnpFramework_UploadFileToFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathLocal = @"C:\Temporary\TestText.txt";
    string fileName = "TestText.txt";
    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
    Microsoft.SharePoint.Client.File myFile =
                                    myFolder.UploadFile(fileName, pathLocal, true);
}
//gavdcodeend 018

//gavdcodebegin 019
static void CsSpPnpFramework_DownloadFileToFolder()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathLocal = @"C:\Temporary\TestText.txt";
    string spFileName = "TestText.txt";
    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
    Microsoft.SharePoint.Client.File myFile = myFolder.GetFile(spFileName);

    ClientResult<System.IO.Stream> myStream = myFile.OpenBinaryStream();
    spPnpCtx.ExecuteQueryRetry();
    using System.IO.FileStream fileStream = new(
                pathLocal, System.IO.FileMode.Create, System.IO.FileAccess.Write);
    myStream.Value.CopyTo(fileStream);
}
//gavdcodeend 019

//gavdcodebegin 020
static void CsSpPnpFramework_CreateItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
    ListItemCreationInformation myInfo = new();
    ListItem newListItem = myList.AddItem(myInfo);
    newListItem["Title"] = "NewListItemPnPFramework";

    newListItem.Update();
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 020

//gavdcodebegin 021
static void CsSpPnpFramework_EnumerateItems()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
    ListItemCollection myListItems = myList.GetItems(CamlQuery.CreateAllItemsQuery());
    spPnpCtx.Load(myListItems);
    spPnpCtx.ExecuteQuery();

    foreach (ListItem oneItem in myListItems)
    {
        Console.WriteLine(oneItem.Id + " - " + oneItem["Title"]);
    }
}
//gavdcodeend 021

//gavdcodebegin 022
static void CsSpPnpFramework_UpdateOneItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
    ListItem myListItem = myList.GetItemById(14);
    myListItem["Title"] = "TitleUpdated";

    myListItem.Update();
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 022

//gavdcodebegin 023
static void CsSpPnpFramework_DeleteOneItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
    ListItem myListItem = myList.GetItemById(14);
    myListItem.DeleteObject();

    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 023

//gavdcodebegin 024
static void CsSpPnpFramework_UploadOneDocument()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string pathLocal = @"C:\Temporary\TestText.txt";
    string fileName = "TestText.txt";
    List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

    using FileStream myFileStream = new(pathLocal, FileMode.Open);
    FileCreationInformation myFileCreationInfo = new()
    {
        Overwrite = true,
        ContentStream = myFileStream,
        Url = fileName
    };
    Microsoft.SharePoint.Client.File newFile =
                            myList.RootFolder.Files.Add(myFileCreationInfo);

    spPnpCtx.Load(newFile);
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 024

//gavdcodebegin 025
static void CsSpPnpFramework_CreateOneAttachment()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    string myFilePath = @"C:\Temporary\TestDocument.docx";
    var myAttachmentInfo = new AttachmentCreationInformation
    {
        FileName = Path.GetFileName(myFilePath)
    };
    using FileStream myFileStream = new(myFilePath, FileMode.Open);
    myAttachmentInfo.ContentStream = myFileStream;
    Attachment myAttachment = myListItem.AttachmentFiles.Add(myAttachmentInfo);
    spPnpCtx.Load(myAttachment);
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 025

//gavdcodebegin 026
static void CsSpPnpFramework_ReadAllAttachments()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spPnpCtx.Load(allAttachments);
    spPnpCtx.ExecuteQuery();

    foreach (Attachment oneAttachment in allAttachments)
    {
        Console.WriteLine("File Name - " + oneAttachment.FileName);
    }
}
//gavdcodeend 026

//gavdcodebegin 027
static void CsSpPnpFramework_DownloadAllAttachments()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    string filePath = @"C:\Temporary";

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spPnpCtx.Load(allAttachments);
    spPnpCtx.ExecuteQuery();

    foreach (Attachment oneAttachment in allAttachments)
    {
        string fileRef = oneAttachment.ServerRelativeUrl;
        Microsoft.SharePoint.Client.File filetoDownload =
                                               myList.RootFolder.Files.GetByUrl(fileRef);
        spPnpCtx.Load(filetoDownload);
        spPnpCtx.ExecuteQuery();

        ClientResult<Stream> fileStream = filetoDownload.OpenBinaryStream();
        spPnpCtx.ExecuteQuery();

        string fileName = oneAttachment.FileName;
        string localPath = Path.Combine(filePath, fileName);
        using FileStream outputFileStream = new(localPath, FileMode.Create);
        fileStream.Value.CopyTo(outputFileStream);
    }
}
//gavdcodeend 027

//gavdcodebegin 028
static void CsSpPnpFramework_DeleteAllAttachments()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spPnpCtx.Load(allAttachments);
    spPnpCtx.ExecuteQuery();

    foreach (Attachment oneAttachment in allAttachments)
    {
        oneAttachment.DeleteObject();
    }

    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 028

//gavdcodebegin 029
static void CsSpPnpFramework_BreakSecurityInheritanceListItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);
    spPnpCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
    spPnpCtx.ExecuteQuery();

    if (myListItem.HasUniqueRoleAssignments == false)
    {
        myListItem.BreakRoleInheritance(false, true);
    }
    myListItem.Update();
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 029

//gavdcodebegin 030
static void CsSpPnpFramework_ResetSecurityInheritanceListItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);
    spPnpCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
    spPnpCtx.ExecuteQuery();

    if (myListItem.HasUniqueRoleAssignments == true)
    {
        myListItem.ResetRoleInheritance();
    }
    myListItem.Update();
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 030

//gavdcodebegin 031
static void CsSpPnpFramework_AddUserToSecurityRoleInListItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition = new(spPnpCtx)
    {
        myWeb.RoleDefinitions.GetByType(RoleType.Reader)
    };
    myListItem.RoleAssignments.Add(myUser, roleDefinition);

    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 031

//gavdcodebegin 032
static void CsSpPnpFramework_UpdateUserSecurityRoleInListItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition = new(spPnpCtx)
    {
        myWeb.RoleDefinitions.GetByType(RoleType.Contributor)
    };

    RoleAssignment myRoleAssignment = myListItem.RoleAssignments.GetByPrincipal(
                                                                        myUser);
    myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

    myRoleAssignment.Update();
    spPnpCtx.ExecuteQuery();
}
//gavdcodeend 032

//gavdcodebegin 033
static void CsSpPnpFramework_DeleteUserFromSecurityRoleInListItem()
{
    using ClientContext spPnpCtx = CsSpPnpFramework_LoginWithAccPw();
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    myListItem.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

    spPnpCtx.ExecuteQuery();
    spPnpCtx.Dispose();
}
//gavdcodeend 033

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 33 ***

//CsSpPnpFramework_CreatePropertyBag();
//CsSpPnpFramework_ReadPropertyBag();
//CsSpPnpFramework_PropertyBagExists();
//CsSpPnpFramework_PropertyBagIndex();
//CsSpPnpFramework_DeletePropertyBag();
//CsSpPnpFramework_DownloadFile();
//CsSpPnpFramework_DownloadFileAsString();
//CsSpPnpFramework_FindFiles();
//CsSpPnpFramework_RequireUpload();
//CsSpPnpFramework_ResetVersion();
//CsSpPnpFramework_CreateFolder();
//CsSpPnpFramework_EnsureFolder();
//CsSpPnpFramework_CreateSubFolder();
//CsSpPnpFramework_FolderExistsBool();
//CsSpPnpFramework_FolderExistsFolder();
//CsSpPnpFramework_FolderExistsWeb();
//CsSpPnpFramework_CreateSubFolderFromFolder();
//CsSpPnpFramework_UploadFileToFolder();
//CsSpPnpFramework_DownloadFileToFolder();
//CsSpPnpFramework_CreateItem();
//CsSpPnpFramework_EnumerateItems();
//CsSpPnpFramework_UpdateOneItem();
//CsSpPnpFramework_DeleteOneItem();
//CsSpPnpFramework_UploadOneDocument();
//CsSpPnpFramework_CreateOneAttachment();
//CsSpPnpFramework_ReadAllAttachments();
//CsSpPnpFramework_DownloadAllAttachments();
//CsSpPnpFramework_DeleteAllAttachments();
//CsSpPnpFramework_BreakSecurityInheritanceListItem();
//CsSpPnpFramework_ResetSecurityInheritanceListItem();
//CsSpPnpFramework_AddUserToSecurityRoleInListItem();
//CsSpPnpFramework_UpdateUserSecurityRoleInListItem();
//CsSpPnpFramework_DeleteUserFromSecurityRoleInListItem();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

