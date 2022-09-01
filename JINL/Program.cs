using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext LoginPnPFramework_WithAccPw()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithCertificate()
{
    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_PnPManagementShell()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithSecret()  //*** LEGACY CODE ***
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

//gavdcodebegin 01
static void SpCsPnPFramework_CreatePropertyBag()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        myList.SetPropertyBagValue("myKey", "myValueString");
    }
}
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsPnPFramework_ReadPropertyBag()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        string myKeyValue = myList.GetPropertyBagValueString("myKey", "");
        Console.WriteLine(myKeyValue);
    }
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsPnPFramework_PropertyBagExists()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        bool myKeyExists = myList.PropertyBagContainsKey("myKey");
        Console.WriteLine(myKeyExists.ToString());
    }
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsPnPFramework_PropertyBagIndex()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsPnPFramework_DeletePropertyBag()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        myList.RemovePropertyBagValue("myKey");
    }
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsPnPFramework_DownloadFile()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathRelative =
                "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";
        string pathLocal = @"C:\Temporary";
        string fileName = "TestText.txt";

        spPnpCtx.Web.SaveFileToLocal(pathRelative, pathLocal, fileName);
    }
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsPnPFramework_DownloadFileAsString()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathRelative =
                "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";
        string myFileAsText = spPnpCtx.Web.GetFileAsString(pathRelative);

        Console.WriteLine(myFileAsText);
    }
}
//gavdcodeend 07

//gavdcodebegin 08
static void SpCsPnPFramework_FindFiles()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        // Very slow method
        List<Microsoft.SharePoint.Client.File> allFiles = spPnpCtx.Web.FindFiles("*.txt");

        foreach (Microsoft.SharePoint.Client.File oneFile in allFiles)
        {
            Console.WriteLine(oneFile.Name + " - " + oneFile.ServerRelativeUrl);
        }
    }
}
//gavdcodeend 08

//gavdcodebegin 09
static void SpCsPnPFramework_RequireUpload()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List<Microsoft.SharePoint.Client.File> myFiles = 
                                                spPnpCtx.Web.FindFiles("TestText.txt");
        Microsoft.SharePoint.Client.File myFile = myFiles[0];

        // Very slow method
        bool requireUpload = myFile.VerifyIfUploadRequired(@"C:\Temporary\TestText.txt");

        Console.WriteLine("Require Upload " + requireUpload.ToString());
    }
}
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsPnPFramework_ResetVersion()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathRelative =
                "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";

        spPnpCtx.Web.ResetFileToPreviousVersion(pathRelative,
                                    CheckinType.MinorCheckIn, "Done by PnPCore");
    }
}
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsPnPFramework_CreateFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.CreateFolder("PnPFrameworkFolder");
    }
}
//gavdcodeend 11

//gavdcodebegin 12
static void SpCsPnPFramework_EnsureFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkEnsureFolder");
    }
}
//gavdcodeend 12

//gavdcodebegin 13
static void SpCsPnPFramework_CreateSubFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
        Folder mySubFolder = myFolder.EnsureFolder("PnPFrameworkSubFolder");
    }
}
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsPnPFramework_FolderExistsBool()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        bool fldExists = myList.RootFolder.FolderExists("PnPFrameworkFolder");
        Console.WriteLine("Folder exists - " + fldExists.ToString());
    }
}
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsPnPFramework_FolderExistsFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder fldExists = myList.RootFolder.ResolveSubFolder("PnPFrameworkFolder");
        Console.WriteLine("Folder exists - " + fldExists.ServerRelativeUrl);
    }
}
//gavdcodeend 15

//gavdcodebegin 16
static void SpCsPnPFramework_FolderExistsWeb()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathRelative =
                "/sites/[SiteName]/[LibraryName]/[FolderName]";

        bool fldExists = spPnpCtx.Site.RootWeb.DoesFolderExists(pathRelative);

        Console.WriteLine("Folder exists - " + fldExists.ToString());
    }
}
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsPnPFramework_CreateSubFolderFromFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.ResolveSubFolder("PnPFrameworkFolder");
        Folder mySubFolder = myFolder.CreateFolder("PnPFrameworkSubFolder02");
    }
}
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsPnPFramework_UploadFileToFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathLocal = @"C:\Temporary\TestText.txt";
        string fileName = "TestText.txt";
        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
        Microsoft.SharePoint.Client.File myFile = 
                                        myFolder.UploadFile(fileName, pathLocal, true);
    }
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsPnPFramework_DownloadFileToFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathLocal = @"C:\Temporary\TestText.txt";
        string spFileName = "TestText.txt";
        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkFolder");
        Microsoft.SharePoint.Client.File myFile = myFolder.GetFile(spFileName);

        ClientResult<System.IO.Stream> myStream = myFile.OpenBinaryStream();
        spPnpCtx.ExecuteQueryRetry();
        using (System.IO.FileStream fileStream = new System.IO.FileStream(
                    pathLocal, System.IO.FileMode.Create, System.IO.FileAccess.Write))
        {
            myStream.Value.CopyTo(fileStream);
        }
    }
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsPnPFramework_CreateItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
        ListItemCreationInformation myInfo = new ListItemCreationInformation();
        ListItem newListItem = myList.AddItem(myInfo);
        newListItem["Title"] = "NewListItemPnPFramework";

        newListItem.Update();
        spPnpCtx.ExecuteQuery();
    }
}
//gavdcodeend 20

//gavdcodebegin 21
static void SpCsPnPFramework_EnumerateItems()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsPnPFramework_UpdateOneItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
        ListItem myListItem = myList.GetItemById(14);
        myListItem["Title"] = "TitleUpdated";

        myListItem.Update();
        spPnpCtx.ExecuteQuery();
    }
}
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsPnPFramework_DeleteOneItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestList");
        ListItem myListItem = myList.GetItemById(14);
        myListItem.DeleteObject();

        spPnpCtx.ExecuteQuery();
    }
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsPnPFramework_UploadOneDocument()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string pathLocal = @"C:\Temporary\TestText.txt";
        string fileName = "TestText.txt";
        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        using (FileStream myFileStream = new FileStream(pathLocal, FileMode.Open))
        {
            FileCreationInformation myFileCreationInfo = new FileCreationInformation
            {
                Overwrite = true,
                ContentStream = myFileStream,
                Url = fileName
            };
            Microsoft.SharePoint.Client.File newFile = myList.RootFolder.Files.Add(myFileCreationInfo);

            spPnpCtx.Load(newFile);
            spPnpCtx.ExecuteQuery();
        }
    }
}
//gavdcodeend 24

//gavdcodebegin 25
static void SpCsPnPFramework_CreateOneAttachment()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");
        int listItemId = 13;
        ListItem myListItem = myList.GetItemById(listItemId);

        string myFilePath = @"C:\Temporary\TestDocument.docx";
        var myAttachmentInfo = new AttachmentCreationInformation();
        myAttachmentInfo.FileName = Path.GetFileName(myFilePath);
        using (FileStream myFileStream = new FileStream(myFilePath, FileMode.Open))
        {
            myAttachmentInfo.ContentStream = myFileStream;
            Attachment myAttachment = myListItem.AttachmentFiles.Add(myAttachmentInfo);
            spPnpCtx.Load(myAttachment);
            spPnpCtx.ExecuteQuery();
        }
    }
}
//gavdcodeend 25

//gavdcodebegin 26
static void SpCsPnPFramework_ReadAllAttachments()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 26

//gavdcodebegin 27
static void SpCsPnPFramework_DownloadAllAttachments()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
            using (FileStream outputFileStream = new FileStream(localPath, FileMode.Create))
            {
                fileStream.Value.CopyTo(outputFileStream);
            }
        }
    }
}
//gavdcodeend 27

//gavdcodebegin 28
static void SpCsPnPFramework_DeleteAllAttachments()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 28

//gavdcodebegin 29
static void SpCsPnPFramework_BreakSecurityInheritanceListItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 29

//gavdcodebegin 30
static void SpCsPnPFramework_ResetSecurityInheritanceListItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
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
}
//gavdcodeend 30

//gavdcodebegin 31
static void SpCsPnPFramework_AddUserToSecurityRoleInListItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("TestList");
        ListItem myListItem = myList.GetItemById(13);

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        RoleDefinitionBindingCollection roleDefinition =
                new RoleDefinitionBindingCollection(spPnpCtx);
        roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
        myListItem.RoleAssignments.Add(myUser, roleDefinition);

        spPnpCtx.ExecuteQuery();
    }
}
//gavdcodeend 31

//gavdcodebegin 32
static void SpCsPnPFramework_UpdateUserSecurityRoleInListItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("TestList");
        ListItem myListItem = myList.GetItemById(13);

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        RoleDefinitionBindingCollection roleDefinition =
                new RoleDefinitionBindingCollection(spPnpCtx);
        roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Contributor));

        RoleAssignment myRoleAssignment = myListItem.RoleAssignments.GetByPrincipal(
                                                                            myUser);
        myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

        myRoleAssignment.Update();
        spPnpCtx.ExecuteQuery();
    }
}
//gavdcodeend 32

//gavdcodebegin 33
static void SpCsPnPFramework_DeleteUserFromSecurityRoleInListItem()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("TestList");
        ListItem myListItem = myList.GetItemById(13);

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        myListItem.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

        spPnpCtx.ExecuteQuery();
        spPnpCtx.Dispose();
    }
}
//gavdcodeend 33

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsPnPFramework_CreatePropertyBag();
//SpCsPnPFramework_ReadPropertyBag();
//SpCsPnPFramework_PropertyBagExists();
//SpCsPnPFramework_PropertyBagIndex();
//SpCsPnPFramework_DeletePropertyBag();
//SpCsPnPFramework_DownloadFile();
//SpCsPnPFramework_DownloadFileAsString();
//SpCsPnPFramework_FindFiles();
//SpCsPnPFramework_RequireUpload();
//SpCsPnPFramework_ResetVersion();
//SpCsPnPFramework_CreateFolder();
//SpCsPnPFramework_EnsureFolder();
//SpCsPnPFramework_CreateSubFolder();
//SpCsPnPFramework_FolderExistsBool();
//SpCsPnPFramework_FolderExistsFolder();
//SpCsPnPFramework_FolderExistsWeb();
//SpCsPnPFramework_CreateSubFolderFromFolder();
//SpCsPnPFramework_UploadFileToFolder();
//SpCsPnPFramework_DownloadFileToFolder();
//SpCsPnPFramework_CreateItem();
//SpCsPnPFramework_EnumerateItems();
//SpCsPnPFramework_UpdateOneItem();
//SpCsPnPFramework_DeleteOneItem();
//SpCsPnPFramework_UploadOneDocument();
//SpCsPnPFramework_CreateOneAttachment();
//SpCsPnPFramework_ReadAllAttachments();
//SpCsPnPFramework_DownloadAllAttachments();
//SpCsPnPFramework_DeleteAllAttachments();
//SpCsPnPFramework_BreakSecurityInheritanceListItem();
//SpCsPnPFramework_ResetSecurityInheritanceListItem();
//SpCsPnPFramework_AddUserToSecurityRoleInListItem();
//SpCsPnPFramework_UpdateUserSecurityRoleInListItem();
//SpCsPnPFramework_DeleteUserFromSecurityRoleInListItem();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable

