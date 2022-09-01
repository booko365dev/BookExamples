using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
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

static PnPContext CreateContextWithInteraction(string TenantId, string ClientId,
                                                   string SiteCollUrl, LogLevel ShowLogs)
{
    IHost myHost = Host.CreateDefaultBuilder()
        .ConfigureServices((context, services) =>
        {
            services.AddPnPCore(options =>
            {
                options.DefaultAuthenticationProvider =
                                    new InteractiveAuthenticationProvider(ClientId,
                                    TenantId,
                                    new Uri("http://localhost"));
            });
        })
        .ConfigureLogging((hostingContext, logging) =>
        {
            logging.SetMinimumLevel(ShowLogs);
        })
        .UseConsoleLifetime()   // Listens for Ctrl+C (Windows) or SIGTERM (Linux)
        .Build();

    myHost.Start();

    IServiceScope myScope = myHost.Services.CreateScope();
    IPnPContextFactory myPnpContextFactory = myScope.ServiceProvider
                                                .GetRequiredService<IPnPContextFactory>();
    Uri mySiteCollUri = new Uri(SiteCollUrl);
    PnPContext myContext = myPnpContextFactory.CreateAsync(mySiteCollUri).Result;

    myHost.Dispose();

    return myContext;
}

static PnPContext CreateContextWithAccPw(string TenantId, string ClientId,
                  string UserAcc, string UserPw, string SiteCollUrl, LogLevel ShowLogs)
{
    IHost myHost = Host.CreateDefaultBuilder()
        .ConfigureServices((context, services) =>
        {
            services.AddPnPCore(options =>
            {
                SecureString secPw = new SecureString();
                foreach (char oneChar in UserPw)
                    secPw.AppendChar(oneChar);

                options.DefaultAuthenticationProvider =
                                    new UsernamePasswordAuthenticationProvider(ClientId,
                                    TenantId,
                                    UserAcc, secPw);
            });
        })
        .ConfigureLogging((hostingContext, logging) =>
        {
            logging.SetMinimumLevel(ShowLogs);
        })
        .UseConsoleLifetime()   // Listens for Ctrl+C (Windows) or SIGTERM (Linux)
        .Build();

    myHost.Start();

    IServiceScope myScope = myHost.Services.CreateScope();
    IPnPContextFactory myPnpContextFactory = myScope.ServiceProvider
                                                .GetRequiredService<IPnPContextFactory>();
    PnPContext myContext = myPnpContextFactory.CreateAsync(new Uri(SiteCollUrl)).Result;

    myHost.Dispose();

    return myContext;
}

static PnPContext CreateContextWithCertificate(string TenantId, string ClientId,
                    string CertificateThumbprint, string SiteCollUrl, LogLevel ShowLogs)
{
    IHost myHost = Host.CreateDefaultBuilder()
        .ConfigureServices((context, services) =>
        {
            services.AddPnPCore(options =>
            {
                options.DefaultAuthenticationProvider =
                                    new X509CertificateAuthenticationProvider(ClientId,
                                    TenantId,
                                    StoreName.My, StoreLocation.CurrentUser,
                                    CertificateThumbprint);
            });
        })
        .ConfigureLogging((hostingContext, logging) =>
        {
            logging.SetMinimumLevel(ShowLogs);
        })
        .UseConsoleLifetime()   // Listens for Ctrl+C (Windows) or SIGTERM (Linux)
        .Build();

    myHost.Start();

    IServiceScope myScope = myHost.Services.CreateScope();
    IPnPContextFactory myPnpContextFactory = myScope.ServiceProvider
                                                .GetRequiredService<IPnPContextFactory>();
    PnPContext myContext = myPnpContextFactory.CreateAsync(new Uri(SiteCollUrl)).Result;

    myHost.Dispose();

    return myContext;
}

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static void SpCsPnPCoreSdk_GetItemsDocuments()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);

        foreach (IListItem oneItem in myList.Items.AsRequested())
        {
            string itemTitle = (oneItem["Title"] != null) ? oneItem["Title"].ToString(): "";
            Console.WriteLine(itemTitle + " - " + 
                              oneItem["ID"].ToString() + " - " +
                              oneItem["FileSystemObjectType"].ToString());
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsPnPCoreSdk_GetItemsByCaml()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string viewXml = @"<View>
                    <ViewFields>
                      <FieldRef Name='Title' />
                      <FieldRef Name='FileRef' />
                    </ViewFields>
                    <Query>
                    </Query>
                   </View>";

        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        myList.LoadItemsByCamlQuery(new CamlQueryOptions()
        {
            ViewXml = viewXml,
            DatesInUtc = true
        });

        foreach (IListItem oneItem in myList.Items.AsRequested())
        {
            string itemTitle = (oneItem["Title"] != null) ? 
                                                    oneItem["Title"].ToString() : "";
            Console.WriteLine(itemTitle + " - " +
                              oneItem["ID"].ToString());
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsPnPCoreSdk_GetOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        string itemTitle = (myItem["Title"] != null) ? myItem["Title"].ToString() : "";
        Console.WriteLine(itemTitle + " - " +
                          myItem["ID"].ToString() + " - " +
                          myItem["FileSystemObjectType"].ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsPnPCoreSdk_GetOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFile myDoc = myContext.Web
            .GetFileByServerRelativeUrl("/sites/[Site]/[Library]/TestText.txt");

        Console.WriteLine(myDoc.Name + " - " + myDoc.UniqueId.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsPnPCoreSdk_GetOneDocumentByFind()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitleAsync("TestLibrary").Result;
        //List<IFile> myFiles = myList.FindFiles("TestText.txt");

        //foreach (IFile oneDoc in myFiles)
        //{
        //    Console.WriteLine(oneDoc.Name + " - " + oneDoc.UniqueId.ToString());
        //}
    }

    Console.WriteLine("Done");
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsPnPCoreSdk_CreateOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList");
        Dictionary<string, object> itemToAdd = new Dictionary<string, object>()
        {
            { "Title", "ItemFromSpCsPnPCoreSdk" },
            { "ColumnText", "This is a text" }
        };

        IListItem newItem = myList.Items.Add(itemToAdd);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsPnPCoreSdk_CreateMultipleItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList");
        for (int itemCounter = 1; itemCounter <= 2; itemCounter++)
        {
            Dictionary<string, object> itemsToAdd = new Dictionary<string, object>
            {
                { "Title", $"NewItem_{itemCounter}" }
            };

            myList.Items.AddBatch(itemsToAdd);
        }

        myContext.Execute();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 07

//gavdcodebegin 08
static void SpCsPnPCoreSdk_UploadOneDocument()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        FileInfo myFileInfo = new FileInfo(filePath);

        FileStream fileToUpload = System.IO.File.OpenRead(myFileInfo.FullName);
        IFolder folderOfLibrary = myContext.Web.Folders
                            .Where(lib => lib.Name == "TestLibrary")
                            .FirstOrDefault();

        IFile addedFile = folderOfLibrary.Files.Add(myFileInfo.Name, fileToUpload);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 08

//gavdcodebegin 09
static void SpCsPnPCoreSdk_DownloadOneDocument()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        string fileUrl = $"{myContext.Uri.PathAndQuery}/TestLibrary/TestText.txt";

        IFile myFile = myContext.Web.GetFileByServerRelativeUrl(fileUrl);

        // Using a Stream
        Stream myFileStream = myFile.GetContent();
        using (var fileStrm = File.Create(filePath))
        {
            myFileStream.Seek(0, SeekOrigin.Begin);
            myFileStream.CopyTo(fileStrm);
        }

        // Using a Byte Array
        byte[] myFileBytes = myFile.GetContentBytes();
        File.WriteAllBytes(filePath, myFileBytes.ToArray());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsPnPCoreSdk_UpdateOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");
        myItem["ColumnText"] = "This is an update";

        myItem.Update();
        //myItem.UpdateOverwriteVersion();
        //myItem.SystemUpdate();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsPnPCoreSdk_UpdateOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFile myDoc = myContext.Web
            .GetFileByServerRelativeUrl("/sites/[Site]/[Library]/TestText.txt");
        myDoc.ListItemAllFields["Title"] = "Document updated";

        myDoc.ListItemAllFields.Update();
        //myDoc.ListItemAllFields.UpdateOverwriteVersion();
        //myDoc.ListItemAllFields.SystemUpdate();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 11

//gavdcodebegin 12
static void SpCsPnPCoreSdk_DeleteOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        myItem.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 12

//gavdcodebegin 13
static void SpCsPnPCoreSdk_DeleteAllItems()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        foreach (var oneItem in myList.Items.AsRequested())
        {
            oneItem.DeleteBatch();
        }

        myContext.Execute();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsPnPCoreSdk_DeleteOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFile myDoc = myContext.Web
            .GetFileByServerRelativeUrl("/sites/[Site]/[Library]/TestText.txt");

        myDoc.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsPnPCoreSdk_GetFolders()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestList",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);

        foreach (IFolder oneFolder in myFolders)
        {
            string itemTitle = (oneFolder.Name != null) ? oneFolder.Name : "";
            Console.WriteLine(itemTitle + " - " +
                              oneFolder.UniqueId);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 15

//gavdcodebegin 16
static void SpCsPnPCoreSdk_GetOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        Console.WriteLine(myFolder.Name + " - " + myFolder.UniqueId);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsPnPCoreSdk_EnsureOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestList",
                                                    p => p.RootFolder).RootFolder;
        bool folderExists = myRootFolder.EnsureFolder("TestFolder").Exists;

        Console.WriteLine("Folder exists -> " + folderExists.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsPnPCoreSdk_AddOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        IFolder newFolder = myRootFolder.Folders.Add("NewFolder");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsPnPCoreSdk_UploadOneFileToOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                            p => p.RootFolder).RootFolder;
        var myFolders = myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        IFile newFileInFolder = myFolder.Files.Add("TestText.txt",
                          System.IO.File.OpenRead(@"C:\Temporary\TestText.txt"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsPnPCoreSdk_DownloadOneFileFromOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        string fileUrl = $"{myContext.Uri.PathAndQuery}/TestLibrary/NewFolder/TestText.txt";

        IFile myFile = myContext.Web.GetFileByServerRelativeUrl(fileUrl);

        // Using a Stream
        Stream myFileStream = myFile.GetContent();
        using (var fileStrm = File.Create(filePath))
        {
            myFileStream.Seek(0, SeekOrigin.Begin);
            myFileStream.CopyTo(fileStrm);
        }

        // Using a Byte Array
        byte[] myFileBytes = myFile.GetContentBytes();
        File.WriteAllBytes(filePath, myFileBytes.ToArray());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 20

//gavdcodebegin 21
static void SpCsPnPCoreSdk_DeleteOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        myFolder.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsPnPCoreSdk_AddAttachmentsToItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        IAttachment addedAttachment = myItem.AttachmentFiles.Add("TestText.txt", 
                System.IO.File.OpenRead(@"C:\Temporary\TestText.txt"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsPnPCoreSdk_GetAttachmentsInItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");
        IAttachmentCollection myAttachments = myItem.AttachmentFiles;

        foreach(IAttachment oneAttachment in myAttachments)
        {
            Console.WriteLine(oneAttachment.FileName);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsPnPCoreSdk_UpdateAttachmentsInItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");
        IAttachmentCollection myAttachments = myItem.AttachmentFiles;

        foreach (IAttachment oneAttachment in myAttachments)
        {
            if (oneAttachment.FileName == "TestText.txt")
            {
                oneAttachment.Delete();
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 24

//gavdcodebegin 25
static void SpCsPnPCoreSdk_BreakeInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        if (myItem != null)
        {
            myItem.BreakRoleInheritance(false, true);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 25

//gavdcodebegin 26
static async void SpCsPnPCoreSdk_HasInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");
        myItem.EnsureProperties(hur => hur.HasUniqueRoleAssignments);

        Console.WriteLine(myItem.HasUniqueRoleAssignments.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 26

//gavdcodebegin 27
static void SpCsPnPCoreSdk_RestoreInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        if (myItem != null)
        {
            myItem.ResetRoleInheritance();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 27

//gavdcodebegin 28
static void SpCsPnPCoreSdk_GetAllSecurityRolesItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        foreach (IRoleAssignment oneRole in myItem.RoleAssignments)
        {
            IRoleDefinitionCollection permLevels = myItem.GetRoleDefinitions(
                                                                oneRole.PrincipalId);

            foreach (IRoleDefinition onePermLevel in permLevels)
            {
                Console.WriteLine(onePermLevel.Name);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 28

//gavdcodebegin 29
static void SpCsPnPCoreSdk_AddSecurityRoleToItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISharePointUser myUser = myContext.Web.GetCurrentUser();
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        myItem.AddRoleDefinitions(myUser.Id, new string[] { "Read", "Edit" });
        myItem.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 29

//gavdcodebegin 30
static void SpCsPnPCoreSdk_DeleteSecurityRoleFromItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISharePointUser myUser = myContext.Web.GetCurrentUser();
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                               .FirstOrDefault(p => p.Title == "ItemOne");

        myItem.RemoveRoleDefinitions(myUser.Id, new string[] { "Read" });
        myItem.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 30


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsPnPCoreSdk_GetItemsDocuments();
//SpCsPnPCoreSdk_GetItemsByCaml();
//SpCsPnPCoreSdk_GetOneItem();
//SpCsPnPCoreSdk_GetOneDocumentByRelative();
//SpCsPnPCoreSdk_GetOneDocumentByFind();
//SpCsPnPCoreSdk_CreateOneItem();
//SpCsPnPCoreSdk_CreateMultipleItem();
//SpCsPnPCoreSdk_UploadOneDocument();
//SpCsPnPCoreSdk_DownloadOneDocument();
//SpCsPnPCoreSdk_UpdateOneItem();
//SpCsPnPCoreSdk_UpdateOneDocumentByRelative();
//SpCsPnPCoreSdk_DeleteOneItem();
//SpCsPnPCoreSdk_DeleteAllItems();
//SpCsPnPCoreSdk_DeleteOneDocumentByRelative();
//SpCsPnPCoreSdk_GetFolders();
//SpCsPnPCoreSdk_GetOneFolder();
//SpCsPnPCoreSdk_EnsureOneFolder();
//SpCsPnPCoreSdk_AddOneFolder();
//SpCsPnPCoreSdk_UploadOneFileToOneFolder();
//SpCsPnPCoreSdk_DownloadOneFileFromOneFolder();
//SpCsPnPCoreSdk_DeleteOneFolder();
//SpCsPnPCoreSdk_AddAttachmentsToItem();
//SpCsPnPCoreSdk_GetAttachmentsInItem();
//SpCsPnPCoreSdk_UpdateAttachmentsInItem();
//SpCsPnPCoreSdk_BreakeInheritanceItem();
//SpCsPnPCoreSdk_HasInheritanceItem();
//SpCsPnPCoreSdk_RestoreInheritanceItem();
//SpCsPnPCoreSdk_GetAllSecurityRolesItem();
//SpCsPnPCoreSdk_AddSecurityRoleToItem();
//SpCsPnPCoreSdk_DeleteSecurityRoleFromItem();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable