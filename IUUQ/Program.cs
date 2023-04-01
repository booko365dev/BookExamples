using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using System.Configuration;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------




//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------




//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 33 ***

SecureString usrPw = new SecureString();
foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
    usrPw.AppendChar(oneChar);

using (AuthenticationManager authenticationManager =
            new AuthenticationManager())
using (ClientContext spCtx = authenticationManager.GetContext(
            new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]),
            ConfigurationManager.AppSettings["UserName"],
            usrPw,
            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
{
    //SpCsCsom_CreateOneItem(spCtx);
    //SpCsCsom_CreateMultipleItem(spCtx);
    //SpCsCsom_UploadOneDocument(spCtx);
    //SpCsCsom_UploadOneDocumentFileStream(spCtx);
    //SpCsCsom_UploadMultipleDocs(spCtx);
    //SpCsCsom_DownloadOneDoc(spCtx);
    //SpCsCsom_DownloadMultipleDocs(spCtx);
    //SpCsCsom_ReadAllListItems(spCtx);
    //SpCsCsom_ReadOneListItem(spCtx);
    //SpCsCsom_ReadAllLibraryDocs(spCtx);
    //SpCsCsom_ReadOneLibraryDoc(spCtx);
    //SpCsCsom_UpdateOneListItem(spCtx);
    //SpCsCsom_UpdateOneLibraryDoc(spCtx);
    //SpCsCsom_DeleteOneListItem(spCtx);
    //SpCsCsom_DeleteAllListItems(spCtx);
    //SpCsCsom_DeleteOneLibraryDoc(spCtx);
    //SpCsCsom_DeleteAllLibraryDocs(spCtx);
    //SpCsCsom_CreateFolderInLibrary(spCtx);
    //SpCsCsom_CreateFolderWithInfo(spCtx);
    //SpCsCsom_AddItemInFolder(spCtx);
    //SpCsCsom_UploadOneDocumentInFolder(spCtx);
    //SpCsCsom_ReadAllFolders(spCtx);
    //SpCsCsom_ReadAllItemsInFolder(spCtx);
    //SpCsCsom_DeleteOneFolder(spCtx);
    //SpCsCsom_CreateOneAttachment(spCtx);
    //SpCsCsom_ReadAllAttachments(spCtx);
    //SpCsCsom_DownloadAllAttachments(spCtx);
    //SpCsCsom_DeleteAllAttachments(spCtx);
    //SpCsCsom_BreakSecurityInheritanceListItem(spCtx);
    //SpCsCsom_ResetSecurityInheritanceListItem(spCtx);
    //SpCsCsom_AddUserToSecurityRoleInListItem(spCtx);
    //SpCsCsom_UpdateUserSecurityRoleInListItem(spCtx);
    //SpCsCsom_DeleteUserFromSecurityRoleInListItem(spCtx);

    Console.WriteLine("Done");
}

//gavdcodebegin 001
static void SpCsCsom_CreateOneItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");

    ListItemCreationInformation myListItemCreationInfo =
                                    new ListItemCreationInformation();
    ListItem newListItem = myList.AddItem(myListItemCreationInfo);
    newListItem["Title"] = "NewListItemCsCsom";

    newListItem.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 001

//gavdcodebegin 012
static void SpCsCsom_CreateMultipleItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");

    for (int intCounter = 0; intCounter < 4; intCounter++)
    {
        ListItemCreationInformation myListItemCreationInfo =
                                        new ListItemCreationInformation();
        ListItem newListItem = myList.AddItem(myListItemCreationInfo);
        newListItem["Title"] = intCounter.ToString() + "-NewListItemCsCsom";
        newListItem.Update();
    }

    spCtx.ExecuteQuery();
}
//gavdcodeend 012

//gavdcodebegin 002
static void SpCsCsom_UploadOneDocument(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

    string filePath = @"C:\Temporary\";
    string fileName = @"TestDocument.docx";

    byte[] myFileContent = System.IO.File.ReadAllBytes(filePath + fileName);
    FileCreationInformation myFileInfo = new FileCreationInformation
    {
        Overwrite = true,
        ContentStream = new MemoryStream(myFileContent),
        Url = fileName
    };

    Microsoft.SharePoint.Client.File uploadFile = myList.RootFolder.Files.Add(myFileInfo);

    spCtx.Load(uploadFile);
    spCtx.ExecuteQuery();
}
//gavdcodeend 002

//gavdcodebegin 022
static void SpCsCsom_UploadOneDocumentFileStream(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

    string filePath = @"C:\Temporary\";
    string fileName = @"TestDocument.docx";

    using (FileStream myFileStream = new
                            FileStream(filePath + fileName, FileMode.Open))
    {
        FileCreationInformation myFileInfo = new FileCreationInformation
        {
            Overwrite = true,
            ContentStream = myFileStream,
            Url = fileName
        };

        Microsoft.SharePoint.Client.File newFile =
                                myList.RootFolder.Files.Add(myFileInfo);
        spCtx.Load(newFile);
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 022

//gavdcodebegin 013
static void SpCsCsom_UploadMultipleDocs(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

    string filesPath = @"C:\Temporary\Docs\";
    string[] myFiles = Directory.GetFiles(filesPath);

    foreach (string oneFile in myFiles)
    {
        using (FileStream myFileStream = new
                                FileStream(oneFile, FileMode.Open))
        {
            FileCreationInformation myFileInfo = new FileCreationInformation
            {
                Overwrite = true,
                ContentStream = myFileStream,
                Url = oneFile.Replace(filesPath, "")
            };

            Microsoft.SharePoint.Client.File newFile =
                                    myList.RootFolder.Files.Add(myFileInfo);

            spCtx.Load(newFile);
            spCtx.ExecuteQuery();
        }
    }
}
//gavdcodeend 013

//gavdcodebegin 003
static void SpCsCsom_DownloadOneDoc(ClientContext spCtx)
{
    //// Using File ID and File GetByUrl
    //List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    //int listItemId = 7;
    //ListItem myListItem = myList.GetItemById(listItemId);
    //spCtx.Load(myListItem);
    //spCtx.Load(myListItem, itm => itm.File);
    //spCtx.ExecuteQuery();
    //string fileRef = myListItem.File.ServerRelativeUrl;
    //Microsoft.SharePoint.Client.File filetoDownload = 
    //                                myList.RootFolder.Files.GetByUrl(fileRef);

    //// Using full File URL
    //string fileUrl = "https://domain.sharepoint.com/sites/sitcoll/library/file.txt";
    //Microsoft.SharePoint.Client.File filetoDownload = spCtx.Web.GetFileByUrl(fileUrl);

    // Using File Relative URL
    string fileRelUrl = "/sites/Test_Guitaca/TestLibrary/TestText.txt";
    Microsoft.SharePoint.Client.File filetoDownload = 
                                        spCtx.Web.GetFileByServerRelativeUrl(fileRelUrl);
    spCtx.Load(filetoDownload);
    spCtx.ExecuteQuery();

    ClientResult<Stream> fileStream = filetoDownload.OpenBinaryStream();
    spCtx.ExecuteQuery();

    string filePath = @"C:\Temporary";
    string fileName = "abc.txt";
    string localPath = Path.Combine(filePath, fileName);
    using (FileStream outputFileStream = new FileStream(localPath, FileMode.Create))
    {
        fileStream.Value.CopyTo(outputFileStream);
    }
}
//gavdcodeend 003

//gavdcodebegin 014
static void SpCsCsom_DownloadMultipleDocs(ClientContext spCtx)
{
    string filePath = @"C:\Temporary";

    FileCollection myFiles = spCtx.Web.GetFolderByServerRelativeUrl(
                                                         "TestLibrary").Files;
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    spCtx.Load(myList);

    spCtx.Load(myFiles);
    spCtx.ExecuteQuery();

    foreach (Microsoft.SharePoint.Client.File oneFile in myFiles)
    {
        string fileRef = oneFile.ServerRelativeUrl;
        Microsoft.SharePoint.Client.File filetoDownload =
                                             myList.RootFolder.Files.GetByUrl(fileRef);
        spCtx.Load(filetoDownload);
        spCtx.ExecuteQuery();

        ClientResult<Stream> fileStream = filetoDownload.OpenBinaryStream();
        spCtx.ExecuteQuery();

        string fileName = oneFile.Name;
        string localPath = Path.Combine(filePath, fileName);
        using (FileStream outputFileStream = new FileStream(localPath, FileMode.Create))
        {
            fileStream.Value.CopyTo(outputFileStream);
        }
    }
}
//gavdcodeend 014

//gavdcodebegin 004
static void SpCsCsom_ReadAllListItems(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItemCollection allItems = myList.GetItems(CamlQuery.CreateAllItemsQuery());
    spCtx.Load(allItems, itms => itms.Include(itm => itm["Title"],
                                             itm => itm.Id));
    spCtx.ExecuteQuery();

    foreach (ListItem oneItem in allItems)
    {
        Console.WriteLine(oneItem["Title"] + " - " + oneItem.Id);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsCsom_ReadOneListItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");

    int filterField = 5;
    int rowLimit = 10;
    string myViewXml = string.Format(@"
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='ID' />
                                <Value Type='Number'>{0}</Value>
                            </Eq>
                        </Where>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='Title' />
                    </ViewFields>
                    <RowLimit>{1}</RowLimit>
                </View>", filterField, rowLimit);

    CamlQuery myCamlQuery = new CamlQuery();
    myCamlQuery.ViewXml = myViewXml;
    ListItemCollection allItems = myList.GetItems(myCamlQuery);
    spCtx.Load(allItems, itms => itms.Include(itm => itm["Title"],
                                             itm => itm.Id));
    spCtx.ExecuteQuery();

    Console.WriteLine("Item Title - " + allItems[0]["Title"]);
}
//gavdcodeend 005

//gavdcodebegin 008
static void SpCsCsom_ReadAllLibraryDocs(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    ListItemCollection allItems = myList.GetItems(CamlQuery.CreateAllItemsQuery());
    spCtx.Load(allItems, itms => itms.Include(itm => itm["FileLeafRef"],
                                             itm => itm.Id));
    spCtx.ExecuteQuery();

    foreach (ListItem oneItem in allItems)
    {
        Console.WriteLine(oneItem["FileLeafRef"] + " - " + oneItem.Id);
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void SpCsCsom_ReadOneLibraryDoc(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

    int filterField = 5;
    int rowLimit = 10;
    string myViewXml = string.Format(@"
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='ID' />
                                <Value Type='Number'>{0}</Value>
                            </Eq>
                        </Where>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='FileLeafRef' />
                    </ViewFields>
                    <RowLimit>{1}</RowLimit>
                </View>", filterField, rowLimit);

    CamlQuery myCamlQuery = new CamlQuery();
    myCamlQuery.ViewXml = myViewXml;
    ListItemCollection allItems = myList.GetItems(myCamlQuery);
    spCtx.Load(allItems, itms => itms.Include(itm => itm["FileLeafRef"],
                                             itm => itm.Id));
    spCtx.ExecuteQuery();

    Console.WriteLine("Item Title - " + allItems[0]["FileLeafRef"]);
}
//gavdcodeend 009

//gavdcodebegin 006
static void SpCsCsom_UpdateOneListItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(6);
    myListItem["Title"] = "NewListItemCsCsomUpdated";

    myListItem.Update();
    spCtx.Load(myListItem);
    spCtx.ExecuteQuery();

    Console.WriteLine("Item Title - " + myListItem["Title"]);
}
//gavdcodeend 006

//gavdcodebegin 010
static void SpCsCsom_UpdateOneLibraryDoc(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    ListItem myListItem = myList.GetItemById(5);
    myListItem["FileLeafRef"] = "LibraryDocCsCsomUpdated.docx";

    myListItem.Update();
    spCtx.Load(myListItem);
    spCtx.ExecuteQuery();

    Console.WriteLine("Item Title - " + myListItem["FileLeafRef"]);
}
//gavdcodeend 010

//gavdcodebegin 007
static void SpCsCsom_DeleteOneListItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(6);
    myListItem.DeleteObject();
    spCtx.ExecuteQuery();
}
//gavdcodeend 007

//gavdcodebegin 015
static void SpCsCsom_DeleteAllListItems(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItemCollection myListItems = myList.GetItems(
                                            CamlQuery.CreateAllItemsQuery());
    spCtx.Load(myListItems);
    spCtx.ExecuteQuery();

    foreach (ListItem oneItem in myListItems)
    {
        ListItem oneItemToDelete = myList.GetItemById(oneItem.Id);
        oneItemToDelete.DeleteObject();
    }

    spCtx.ExecuteQuery();
}
//gavdcodeend 015

//gavdcodebegin 011
static void SpCsCsom_DeleteOneLibraryDoc(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    ListItem myListItem = myList.GetItemById(6);
    myListItem.DeleteObject();
    spCtx.ExecuteQuery();
}
//gavdcodeend 011

//gavdcodebegin 016
static void SpCsCsom_DeleteAllLibraryDocs(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
    ListItemCollection myListItems = myList.GetItems(
                                            CamlQuery.CreateAllItemsQuery());
    spCtx.Load(myListItems);
    spCtx.ExecuteQuery();

    foreach (ListItem oneItem in myListItems)
    {
        ListItem oneItemToDelete = myList.GetItemById(oneItem.Id);
        oneItemToDelete.DeleteObject();
    }

    spCtx.ExecuteQuery();
}
//gavdcodeend 016

//gavdcodebegin 023
static void SpCsCsom_CreateFolderInLibrary(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestLibrary");

    Folder myFolder01 = myList.RootFolder.Folders.Add("FirstLevelFolder");
    myFolder01.Update();
    Folder mySubFolder = myFolder01.Folders.Add("SecondLevelFolder");
    mySubFolder.Update();

    spCtx.ExecuteQuery();
    spCtx.Dispose();
}
//gavdcodeend 023

//gavdcodebegin 024
static void SpCsCsom_CreateFolderWithInfo(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");

    ListItemCreationInformation infoFolder = new ListItemCreationInformation();
    infoFolder.UnderlyingObjectType = FileSystemObjectType.Folder;
    infoFolder.LeafName = "FolderWithInfo";
    ListItem newItem = myList.AddItem(infoFolder);
    newItem["Title"] = "FolderWithInfo";
    newItem.Update();

    spCtx.ExecuteQuery();
    spCtx.Dispose();
}
//gavdcodeend 024

//gavdcodebegin 025
static void SpCsCsom_AddItemInFolder(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");

    ListItemCreationInformation myListItemCreationInfo =
        new ListItemCreationInformation
        {
            FolderUrl = string.Format("{0}/lists/{1}/{2}", spCtx.Url,
                                                "TestList", "FolderWithInfo")
        };
    ListItem newListItem = myList.AddItem(myListItemCreationInfo);
    newListItem["Title"] = "NewListItemInFolderCsCsom";
    newListItem.Update();

    spCtx.ExecuteQuery();
    spCtx.Dispose();
}
//gavdcodeend 025

//gavdcodebegin 026
static void SpCsCsom_UploadOneDocumentInFolder(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

    string filePath = @"C:\Temporary\";
    string fileName = @"TestDocument.docx";

    using (FileStream myFileStream = new
                            FileStream(filePath + fileName, FileMode.Open))
    {
        FileCreationInformation myFileCreationInfo = new FileCreationInformation
        {
            Overwrite = true,
            ContentStream = myFileStream,
            Url = string.Format("{0}/{1}/{2}/{3}", spCtx.Url, "TestLibrary",
                                                "FirstLevelFolder", fileName)
        };

        Microsoft.SharePoint.Client.File newFile =
                                myList.RootFolder.Files.Add(myFileCreationInfo);
        spCtx.Load(newFile);
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 026

//gavdcodebegin 027
static void SpCsCsom_ReadAllFolders(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItemCollection allItems = myList.GetItems(CamlQuery.CreateAllFoldersQuery());
    spCtx.Load(allItems, itms => itms.Include(itm => itm.Folder));

    spCtx.ExecuteQuery();

    List<Folder> allFolders = allItems.Select(itm => itm.Folder).ToList();

    foreach (Folder oneFolder in allFolders)
    {
        Console.WriteLine(oneFolder.Name + " - " + oneFolder.ServerRelativeUrl);
    }
}
//gavdcodeend 027

//gavdcodebegin 028
static void SpCsCsom_ReadAllItemsInFolder(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    CamlQuery myQuery = CamlQuery.CreateAllItemsQuery();
    myQuery.FolderServerRelativeUrl = "/sites/[SiteName]/Lists/TestList/FolderWithInfo";
    ListItemCollection allItems = myList.GetItems(myQuery);
    spCtx.Load(allItems, itms => itms.Include(itm => itm["Title"],
                                             itm => itm.Id));
    spCtx.ExecuteQuery();

    foreach (ListItem oneItem in allItems)
    {
        Console.WriteLine(oneItem["Title"] + " - " + oneItem.Id);
    }
}
//gavdcodeend 028

//gavdcodebegin 029
static void SpCsCsom_DeleteOneFolder(ClientContext spCtx)
{
    string folderRelativeUrl = "/sites/[SiteName]/Lists/TestList/FolderWithInfo";
    Folder myFolder = spCtx.Web.GetFolderByServerRelativeUrl(folderRelativeUrl);

    myFolder.DeleteObject();
    spCtx.ExecuteQuery();
}
//gavdcodeend 029

//gavdcodebegin 030
static void SpCsCsom_CreateOneAttachment(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    string myFilePath = @"C:\Temporary\TestDocument.docx";
    var myAttachmentInfo = new AttachmentCreationInformation();
    myAttachmentInfo.FileName = Path.GetFileName(myFilePath);
    using (FileStream myFileStream = new FileStream(myFilePath, FileMode.Open))
    {
        myAttachmentInfo.ContentStream = myFileStream;
        Attachment myAttachment = myListItem.AttachmentFiles.Add(myAttachmentInfo);
        spCtx.Load(myAttachment);
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 030

//gavdcodebegin 031
static void SpCsCsom_ReadAllAttachments(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spCtx.Load(allAttachments);
    spCtx.ExecuteQuery();

    foreach (Attachment oneAttachment in allAttachments)
    {
        Console.WriteLine("File Name - " + oneAttachment.FileName);
    }
}
//gavdcodeend 031

//gavdcodebegin 032
static void SpCsCsom_DownloadAllAttachments(ClientContext spCtx)
{
    string filePath = @"C:\Temporary";

    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spCtx.Load(allAttachments);
    spCtx.ExecuteQuery();

    string myFilesPath = @"C:\Temporary\";
    foreach (Attachment oneAttachment in allAttachments)
    {
        string fileRef = oneAttachment.ServerRelativeUrl;
        Microsoft.SharePoint.Client.File filetoDownload =
                                           myList.RootFolder.Files.GetByUrl(fileRef);
        spCtx.Load(filetoDownload);
        spCtx.ExecuteQuery();

        ClientResult<Stream> fileStream = filetoDownload.OpenBinaryStream();
        spCtx.ExecuteQuery();

        string fileName = oneAttachment.FileName;
        string localPath = Path.Combine(filePath, fileName);
        using (FileStream outputFileStream = new FileStream(localPath, FileMode.Create))
        {
            fileStream.Value.CopyTo(outputFileStream);
        }
    }
}
//gavdcodeend 032

//gavdcodebegin 033
static void SpCsCsom_DeleteAllAttachments(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    int listItemId = 13;
    ListItem myListItem = myList.GetItemById(listItemId);

    AttachmentCollection allAttachments = myListItem.AttachmentFiles;
    spCtx.Load(allAttachments);
    spCtx.ExecuteQuery();

    foreach (Attachment oneAttachment in allAttachments)
    {
        oneAttachment.DeleteObject();
    }

    spCtx.ExecuteQuery();
}
//gavdcodeend 033

//gavdcodebegin 017
static void SpCsCsom_BreakSecurityInheritanceListItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);
    spCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
    spCtx.ExecuteQuery();

    if (myListItem.HasUniqueRoleAssignments == false)
    {
        myListItem.BreakRoleInheritance(false, true);
    }
    myListItem.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 017

//gavdcodebegin 018
static void SpCsCsom_ResetSecurityInheritanceListItem(ClientContext spCtx)
{
    List myList = spCtx.Web.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);
    spCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
    spCtx.ExecuteQuery();

    if (myListItem.HasUniqueRoleAssignments == true)
    {
        myListItem.ResetRoleInheritance();
    }
    myListItem.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 018

//gavdcodebegin 019
static void SpCsCsom_AddUserToSecurityRoleInListItem(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition =
            new RoleDefinitionBindingCollection(spCtx);
    roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
    myListItem.RoleAssignments.Add(myUser, roleDefinition);

    spCtx.ExecuteQuery();
}
//gavdcodeend 019

//gavdcodebegin 020
static void SpCsCsom_UpdateUserSecurityRoleInListItem(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition =
            new RoleDefinitionBindingCollection(spCtx);
    roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Contributor));

    RoleAssignment myRoleAssignment = myListItem.RoleAssignments.GetByPrincipal(
                                                                        myUser);
    myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

    myRoleAssignment.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 020

//gavdcodebegin 021
static void SpCsCsom_DeleteUserFromSecurityRoleInListItem(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("TestList");
    ListItem myListItem = myList.GetItemById(13);

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    myListItem.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

    spCtx.ExecuteQuery();
    spCtx.Dispose();
}
//gavdcodeend 021

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

public class AuthenticationManager : IDisposable
{
    private static readonly HttpClient httpClient = new HttpClient();
    private const string tokenEndpoint =
                            "https://login.microsoftonline.com/common/oauth2/token";

    private static readonly SemaphoreSlim semaphoreSlimTokens = new SemaphoreSlim(1);
    private AutoResetEvent tokenResetEvent = null;
    private readonly ConcurrentDictionary<string, string> tokenCache =
                                            new ConcurrentDictionary<string, string>();
    private bool disposedValue;

    internal class TokenWaitInfo
    {
        public RegisteredWaitHandle Handle = null;
    }

    public ClientContext GetContext(Uri web, string userPrincipalName,
                                            SecureString userPassword, string clientId)
    {
        var context = new ClientContext(web);

        context.ExecutingWebRequest += (sender, e) =>
        {
            string accessToken = EnsureAccessTokenAsync(
               new Uri($"{web.Scheme}://{web.DnsSafeHost}"),
               userPrincipalName,
               new System.Net.NetworkCredential(string.Empty, userPassword).Password,
               clientId).GetAwaiter().GetResult();

            if (accessToken.Contains("TokenErrorException") == true)
            {
                throw new Exception(accessToken); // An error has been raised by AAD
            }

            e.WebRequestExecutor.RequestHeaders["Authorization"] =
                "Bearer " + accessToken;
        };

        return context;
    }

    public async Task<string> EnsureAccessTokenAsync(Uri resourceUri,
                        string userPrincipalName, string userPassword, string clientId)
    {
        string accessTokenFromCache = TokenFromCache(resourceUri, tokenCache);
        if (accessTokenFromCache == null)
        {
            await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
            try
            {
                string accessToken = await AcquireTokenAsync(resourceUri,
                    userPrincipalName, userPassword, clientId).ConfigureAwait(false);

                if (accessToken.Contains("TokenErrorException") == true)
                { return accessToken; } // An error has been raised by Azure AD

                AddTokenToCache(resourceUri, tokenCache, accessToken);

                tokenResetEvent = new AutoResetEvent(false);
                TokenWaitInfo wi = new TokenWaitInfo();
                wi.Handle = ThreadPool.RegisterWaitForSingleObject(
                    tokenResetEvent,
                    async (state, timedOut) =>
                    {
                        if (!timedOut)
                        {
                            TokenWaitInfo wi1 = (TokenWaitInfo)state;
                            if (wi1.Handle != null)
                            {
                                wi1.Handle.Unregister(null);
                            }
                        }
                        else
                        {
                            try
                            {
                                await semaphoreSlimTokens.WaitAsync().
                                                            ConfigureAwait(false);
                                RemoveTokenFromCache(resourceUri, tokenCache);
                            }
                            catch (Exception ex)
                            {
                                RemoveTokenFromCache(resourceUri, tokenCache);
                            }
                            finally
                            {
                                semaphoreSlimTokens.Release();
                            }
                        }
                    },
                    wi,
                    (uint)CalculateThreadSleep(accessToken).TotalMilliseconds,
                    true
                );

                return accessToken;
            }
            finally
            {
                semaphoreSlimTokens.Release();
            }
        }
        else
        {
            return accessTokenFromCache;
        }
    }

    private async Task<string> AcquireTokenAsync(Uri resourceUri,
                                        string username, string password, string clientId)
    {
        string resource = $"{resourceUri.Scheme}://{resourceUri.DnsSafeHost}";

        var body = $"resource={resource}&";
        body += $"client_id={clientId}&";
        body += $"grant_type=password&";
        body += $"username={HttpUtility.UrlEncode(username)}&";
        body += $"password={HttpUtility.UrlEncode(password)}";
        using (var stringContent = new StringContent(body,
                            Encoding.UTF8, "application/x-www-form-urlencoded"))
        {
            var result = await httpClient.PostAsync(tokenEndpoint,
                            stringContent).ContinueWith((response) =>
                            {
                                return response.Result.Content.ReadAsStringAsync().Result;
                            }).ConfigureAwait(false);

            var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
            try
            { // Check for an error returned by Azure AD
                var tokenError = tokenResult.GetProperty("error").GetString();

                string strError = "TokenErrorException - " +
                            tokenResult.GetProperty("error").GetString() + " - " +
                            tokenResult.GetProperty("error_description").GetString();

                return strError;
            }
            catch
            { } // Nothing to catch, the response is giving correctly the token 

            var token = tokenResult.GetProperty("access_token").GetString();
            return token;
        }
    }

    private static string TokenFromCache(Uri web, ConcurrentDictionary<string,
                                                                    string> tokenCache)
    {
        if (tokenCache.TryGetValue(web.DnsSafeHost, out string accessToken))
        {
            return accessToken;
        }

        return null;
    }

    private static void AddTokenToCache(Uri web, ConcurrentDictionary<string,
                                            string> tokenCache, string newAccessToken)
    {
        if (tokenCache.TryGetValue(web.DnsSafeHost, out string currentAccessToken))
        {
            tokenCache.TryUpdate(web.DnsSafeHost, newAccessToken, currentAccessToken);
        }
        else
        {
            tokenCache.TryAdd(web.DnsSafeHost, newAccessToken);
        }
    }

    private static void RemoveTokenFromCache(Uri web, ConcurrentDictionary<string,
                                                                    string> tokenCache)
    {
        tokenCache.TryRemove(web.DnsSafeHost, out string currentAccessToken);
    }

    private static TimeSpan CalculateThreadSleep(string accessToken)
    {
        var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
        var lease = GetAccessTokenLease(token.ValidTo);
        lease = TimeSpan.FromSeconds(lease.TotalSeconds -
            TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds -
            TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
        return lease;
    }

    private static TimeSpan GetAccessTokenLease(DateTime expiresOn)
    {
        DateTime now = DateTime.UtcNow;
        DateTime expires = expiresOn.Kind == DateTimeKind.Utc ? expiresOn :
            TimeZoneInfo.ConvertTimeToUtc(expiresOn);
        TimeSpan lease = expires - now;
        return lease;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                if (tokenResetEvent != null)
                {
                    tokenResetEvent.Set();
                    tokenResetEvent.Dispose();
                }
            }

            disposedValue = true;
        }
    }

    // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method  
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}

#nullable enable
