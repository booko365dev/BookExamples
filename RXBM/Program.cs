﻿using Microsoft.Extensions.DependencyInjection;
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
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static PnPContext CsPsPnpCoreSdk_CreateContextWithInteraction(
    string TenantId, string ClientId, string SiteCollUrl, LogLevel ShowLogs)
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
    Uri mySiteCollUri = new(SiteCollUrl);
    PnPContext myContext = myPnpContextFactory.CreateAsync(mySiteCollUri).Result;

    myHost.Dispose();

    return myContext;
}

static PnPContext CsPsPnpCoreSdk_CreateContextWithAccPw(
    string TenantId, string ClientId, string UserAcc, string UserPw, 
    string SiteCollUrl, LogLevel ShowLogs)
{
    IHost myHost = Host.CreateDefaultBuilder()
        .ConfigureServices((context, services) =>
        {
            services.AddPnPCore(options =>
            {
                SecureString secPw = new();
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

static PnPContext CsPsPnpCoreSdk_CreateContextWithCertificate(
    string TenantId, string ClientId, string CertificateThumbprint, 
    string SiteCollUrl, LogLevel ShowLogs)
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

//gavdcodebegin 001
static void CsSpPnpCoreSdk_GetItemsDocuments()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);

        foreach (IListItem oneItem in myList.Items.AsRequested())
        {
            string itemTitle = (oneItem["Title"] != null) ? 
                                        oneItem["Title"].ToString(): "";
            Console.WriteLine(itemTitle + " - " + 
                              oneItem["ID"].ToString() + " - " +
                              oneItem["FileSystemObjectType"].ToString());
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnpCoreSdk_GetItemsByCaml()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnpCoreSdk_GetOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnpCoreSdk_GetOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFile myDoc = myContext.Web
            .GetFileByServerRelativeUrl("/sites/[Site]/[Library]/TestText.txt");

        Console.WriteLine(myDoc.Name + " - " + myDoc.UniqueId.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnpCoreSdk_GetOneDocumentByFind()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnpCoreSdk_CreateOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList");
        Dictionary<string, object> itemToAdd = new()
        {
            { "Title", "ItemFromSpCsPnPCoreSdk" },
            { "ColumnText", "This is a text" }
        };

        IListItem newItem = myList.Items.Add(itemToAdd);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnpCoreSdk_CreateMultipleItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList");
        for (int itemCounter = 1; itemCounter <= 2; itemCounter++)
        {
            Dictionary<string, object> itemsToAdd = new()
            {
                { "Title", $"NewItem_{itemCounter}" }
            };

            myList.Items.AddBatch(itemsToAdd);
        }

        myContext.Execute();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnpCoreSdk_UploadOneDocument()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        FileInfo myFileInfo = new(filePath);

        FileStream fileToUpload = System.IO.File.OpenRead(myFileInfo.FullName);
        IFolder folderOfLibrary = myContext.Web.Folders
                            .Where(lib => lib.Name == "TestLibrary")
                            .FirstOrDefault();

        IFile addedFile = folderOfLibrary.Files.Add(myFileInfo.Name, fileToUpload);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpPnpCoreSdk_DownloadOneDocument()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        string fileUrl = $"{myContext.Uri.PathAndQuery}/TestLibrary/TestText.txt";

        IFile myFile = myContext.Web.GetFileByServerRelativeUrl(fileUrl);

        // Using a Stream
        Stream myFileStream = myFile.GetContent();
        using (FileStream fileStream = File.Create(filePath))
        {
            myFileStream.Seek(0, SeekOrigin.Begin);
            myFileStream.CopyTo(fileStream);
        }

        // Using a Byte Array
        byte[] myFileBytes = myFile.GetContentBytes();
        File.WriteAllBytes(filePath, myFileBytes.ToArray());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpPnpCoreSdk_UpdateOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpPnpCoreSdk_UpdateOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpPnpCoreSdk_DeleteOneItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                          .FirstOrDefault(p => p.Title == "ItemOne");

        myItem.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpPnpCoreSdk_DeleteAllItems()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpPnpCoreSdk_DeleteOneDocumentByRelative()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFile myDoc = myContext.Web
            .GetFileByServerRelativeUrl("/sites/[Site]/[Library]/TestText.txt");

        myDoc.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpPnpCoreSdk_GetFolders()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestList",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = 
                myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);

        foreach (IFolder oneFolder in myFolders)
        {
            string itemTitle = (oneFolder.Name != null) ? oneFolder.Name : "";
            Console.WriteLine(itemTitle + " - " + oneFolder.UniqueId);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpPnpCoreSdk_GetOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = 
                myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = 
                myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        Console.WriteLine(myFolder.Name + " - " + myFolder.UniqueId);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsSpPnpCoreSdk_EnsureOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestList",
                                                    p => p.RootFolder).RootFolder;
        bool folderExists = myRootFolder.EnsureFolder("TestFolder").Exists;

        Console.WriteLine("Folder exists -> " + folderExists.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 017

//gavdcodebegin 018
static void CsSpPnpCoreSdk_AddOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        IFolder newFolder = myRootFolder.Folders.Add("NewFolder");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 018

//gavdcodebegin 019
static void CsSpPnpCoreSdk_UploadOneFileToOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                            p => p.RootFolder).RootFolder;
        var myFolders = 
                myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = 
                myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        IFile newFileInFolder = myFolder.Files.Add("TestText.txt",
                          System.IO.File.OpenRead(@"C:\Temporary\TestText.txt"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 019

//gavdcodebegin 020
static void CsSpPnpCoreSdk_DownloadOneFileFromOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string filePath = @"C:\Temporary\TestText.txt";
        string fileUrl = 
                  $"{myContext.Uri.PathAndQuery}/TestLibrary/NewFolder/TestText.txt";

        IFile myFile = myContext.Web.GetFileByServerRelativeUrl(fileUrl);

        // Using a Stream
        Stream myFileStream = myFile.GetContent();
        using (var fileStream = File.Create(filePath))
        {
            myFileStream.Seek(0, SeekOrigin.Begin);
            myFileStream.CopyTo(fileStream);
        }

        // Using a Byte Array
        byte[] myFileBytes = myFile.GetContentBytes();
        File.WriteAllBytes(filePath, myFileBytes.ToArray());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 020

//gavdcodebegin 021
static void CsSpPnpCoreSdk_DeleteOneFolder()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFolder myRootFolder = myContext.Web.Lists.GetByTitle("TestLibrary",
                                                    p => p.RootFolder).RootFolder;
        var myFolders = 
                myRootFolder.Folders.QueryProperties(p => p.Folders, p => p.Name);
        IFolder myFolder = 
                myFolders.Where(f => f.Name == "NewFolder").FirstOrDefault();

        myFolder.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 021

//gavdcodebegin 022
static void CsSpPnpCoreSdk_AddAttachmentsToItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                          .FirstOrDefault(p => p.Title == "ItemOne");

        IAttachment addedAttachment = myItem.AttachmentFiles.Add("TestText.txt", 
                System.IO.File.OpenRead(@"C:\Temporary\TestText.txt"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 022

//gavdcodebegin 023
static void CsSpPnpCoreSdk_GetAttachmentsInItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 023

//gavdcodebegin 024
static void CsSpPnpCoreSdk_UpdateAttachmentsInItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 024

//gavdcodebegin 025
static void CsSpPnpCoreSdk_BreakInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 025

//gavdcodebegin 026
static void CsSpPnpCoreSdk_HasInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("TestList", p => p.Items);
        IListItem myItem = myList.Items.AsRequested()
                                          .FirstOrDefault(p => p.Title == "ItemOne");
        myItem.EnsureProperties(hur => hur.HasUniqueRoleAssignments);

        Console.WriteLine(myItem.HasUniqueRoleAssignments.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 026

//gavdcodebegin 027
static void CsSpPnpCoreSdk_RestoreInheritanceItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 027

//gavdcodebegin 028
static void CsSpPnpCoreSdk_GetAllSecurityRolesItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 028

//gavdcodebegin 029
static void CsSpPnpCoreSdk_AddSecurityRoleToItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 029

//gavdcodebegin 030
static void CsSpPnpCoreSdk_DeleteSecurityRoleFromItem()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPsPnpCoreSdk_CreateContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 030


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 30 ***

//CsSpPnpCoreSdk_GetItemsDocuments();
//CsSpPnpCoreSdk_GetItemsByCaml();
//CsSpPnpCoreSdk_GetOneItem();
//CsSpPnpCoreSdk_GetOneDocumentByRelative();
//CsSpPnpCoreSdk_GetOneDocumentByFind();
//CsSpPnpCoreSdk_CreateOneItem();
//CsSpPnpCoreSdk_CreateMultipleItem();
//CsSpPnpCoreSdk_UploadOneDocument();
//CsSpPnpCoreSdk_DownloadOneDocument();
//CsSpPnpCoreSdk_UpdateOneItem();
//CsSpPnpCoreSdk_UpdateOneDocumentByRelative();
//CsSpPnpCoreSdk_DeleteOneItem();
//CsSpPnpCoreSdk_DeleteAllItems();
//CsSpPnpCoreSdk_DeleteOneDocumentByRelative();
//CsSpPnpCoreSdk_GetFolders();
//CsSpPnpCoreSdk_GetOneFolder();
//CsSpPnpCoreSdk_EnsureOneFolder();
//CsSpPnpCoreSdk_AddOneFolder();
//CsSpPnpCoreSdk_UploadOneFileToOneFolder();
//CsSpPnpCoreSdk_DownloadOneFileFromOneFolder();
//CsSpPnpCoreSdk_DeleteOneFolder();
//CsSpPnpCoreSdk_AddAttachmentsToItem();
//CsSpPnpCoreSdk_GetAttachmentsInItem();
//CsSpPnpCoreSdk_UpdateAttachmentsInItem();
//CsSpPnpCoreSdk_BreakInheritanceItem();
//CsSpPnpCoreSdk_HasInheritanceItem();
//CsSpPnpCoreSdk_RestoreInheritanceItem();
//CsSpPnpCoreSdk_GetAllSecurityRolesItem();
//CsSpPnpCoreSdk_AddSecurityRoleToItem();
//CsSpPnpCoreSdk_DeleteSecurityRoleFromItem();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

