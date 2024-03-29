﻿using Microsoft.SharePoint.Client;
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

//gavdcodebegin 001
static void SpCsPnPFramework_CreatePropertyBag()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        myList.SetPropertyBagValue("myKey", "myValueString");
    }
}
//gavdcodeend 001

//gavdcodebegin 002
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
//gavdcodeend 002

//gavdcodebegin 003
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
//gavdcodeend 003

//gavdcodebegin 004
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
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsPnPFramework_DeletePropertyBag()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Web.Lists.GetByTitle("TestList");

        myList.RemovePropertyBagValue("myKey");
    }
}
//gavdcodeend 005

//gavdcodebegin 006
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
//gavdcodeend 006

//gavdcodebegin 007
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
//gavdcodeend 007

//gavdcodebegin 008
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
//gavdcodeend 008

//gavdcodebegin 009
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
//gavdcodeend 009

//gavdcodebegin 010
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
//gavdcodeend 010

//gavdcodebegin 011
static void SpCsPnPFramework_CreateFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.CreateFolder("PnPFrameworkFolder");
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void SpCsPnPFramework_EnsureFolder()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List myList = spPnpCtx.Site.RootWeb.GetListByTitle("TestLibrary");

        Folder myFolder = myList.RootFolder.EnsureFolder("PnPFrameworkEnsureFolder");
    }
}
//gavdcodeend 012

//gavdcodebegin 013
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
//gavdcodeend 013

//gavdcodebegin 014
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
//gavdcodeend 014

//gavdcodebegin 015
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
//gavdcodeend 015

//gavdcodebegin 016
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
//gavdcodeend 016

//gavdcodebegin 017
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
//gavdcodeend 017

//gavdcodebegin 018
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
//gavdcodeend 018

//gavdcodebegin 019
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
//gavdcodeend 019

//gavdcodebegin 020
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
//gavdcodeend 020

//gavdcodebegin 021
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
//gavdcodeend 021

//gavdcodebegin 022
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
//gavdcodeend 022

//gavdcodebegin 023
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
//gavdcodeend 023

//gavdcodebegin 024
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
            Microsoft.SharePoint.Client.File newFile = 
                                    myList.RootFolder.Files.Add(myFileCreationInfo);

            spPnpCtx.Load(newFile);
            spPnpCtx.ExecuteQuery();
        }
    }
}
//gavdcodeend 024

//gavdcodebegin 025
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
//gavdcodeend 025

//gavdcodebegin 026
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
//gavdcodeend 026

//gavdcodebegin 027
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
//gavdcodeend 027

//gavdcodebegin 028
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
//gavdcodeend 028

//gavdcodebegin 029
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
//gavdcodeend 029

//gavdcodebegin 030
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
//gavdcodeend 030

//gavdcodebegin 031
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
//gavdcodeend 031

//gavdcodebegin 032
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
//gavdcodeend 032

//gavdcodebegin 033
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
//gavdcodeend 033

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 33 ***

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

