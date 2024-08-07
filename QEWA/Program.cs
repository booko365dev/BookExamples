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

static PnPContext CsPnpCoreSdk_GetContextWithInteraction(string TenantId, string ClientId,
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
    Uri mySiteCollUri = new(SiteCollUrl);
    PnPContext myContext = myPnpContextFactory.CreateAsync(mySiteCollUri).Result;

    myHost.Dispose();

    return myContext;
}

static PnPContext CsPnpCoreSdk_GetContextWithAccPw(string TenantId, string ClientId,
                  string UserAcc, string UserPw, string SiteCollUrl, LogLevel ShowLogs)
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

static PnPContext CsPnpCoreSdk_GetContextWithCertificate(string TenantId, string ClientId,
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


//gavdcodebegin 001
static void CsSpPnPCoreSdk_GetAllLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IListCollection myLists = myContext.Web.Lists;

        foreach (IList oneList in myLists)
        {
            Console.WriteLine(oneList.Title);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnPCoreSdk_GetOneList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myOneList = myContext.Web.Lists.GetByTitle("Documents");
        //IList myOneList = myContext.Web.Lists.Where(
        //                           lst => lst.Title == "Documents").FirstOrDefault();
        //IList myOneList = myContext.Web.Lists.GetById(
        //                           new Guid("32243fc3-33dc-4b56-bec5-3166206c26ad"));
        //IList myOneList = myContext.Web.Lists.GetByServerRelativeUrl(
        //                           $"{myContext.Uri.PathAndQuery}/Shared Documents");

        Console.WriteLine(myOneList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnPCoreSdk_GetOneListProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("Documents",
                                                      lst => lst.Id,
                                                      lst => lst.TemplateType);

        Console.WriteLine(myList.Id + " - " + myList.TemplateType);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnPCoreSdk_CreateList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.Add(
                                    "NewListPnPCoreSDK", ListTemplateType.GenericList);

        Console.WriteLine(myList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnPCoreSdk_UpdateList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Description = "New Description for List";
        myList.Update();

        Console.WriteLine(myList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnPCoreSdk_RecycleList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Recycle();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnPCoreSdk_DeleteList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnPCoreSdk_GetAllFieldsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IFieldCollection myListFields = myContext.Web.Lists.GetByTitle(
                                                            "NewListPnPCoreSDK").Fields;

        foreach (IField oneField in myListFields)
        {
            Console.WriteLine(oneField.Title);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpPnPCoreSdk_GetPropertiesFieldsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK",
                            lst => lst.Fields.QueryProperties(
                                lst => lst.Id,
                                lst => lst.InternalName,
                                lst => lst.FieldTypeKind));

        foreach (IField oneField in myList.Fields.AsRequested())
        {
            Console.WriteLine(oneField.InternalName + " - " +
                              oneField.Id + " - " +
                              oneField.FieldTypeKind);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpPnPCoreSdk_CreateFieldByXmlToLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myFieldXml =
          @"<Field Type=""Text"" Name=""myTextField"" DisplayName=""My Text Field""/>";

        IField myListField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Fields.AddFieldAsXml(myFieldXml, true);

        Console.WriteLine(myListField.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpPnPCoreSdk_CreateFieldByApiToLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IField myListField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Fields.AddMultilineText("My Multiline Field",
                                                    new FieldMultilineTextOptions()
                                                    {
                                                        Group = "Custom Fields",
                                                        AddToDefaultView = true,
                                                        RichText = true
                                                    });

        Console.WriteLine(myListField.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpPnPCoreSdk_UpdateFieldInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IField myField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            Fields.Where(fld => fld.Title == "My Text Field").FirstOrDefault();

        if (myField != null)
        {
            myField.Description = "New Description Field";
            myField.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpPnPCoreSdk_DeleteFieldFromLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IField myField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            Fields.Where(fld => fld.Title == "My Text Field").FirstOrDefault();

        if (myField != null)
        {
            myField.Delete();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpPnPCoreSdk_GetAllViewsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IViewCollection myListViews = myContext.Web.Lists.GetByTitle(
                                                         "NewListPnPCoreSDK").Views;

        foreach (IView oneView in myListViews)
        {
            Console.WriteLine(oneView.Title);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpPnPCoreSdk_CreateViewForLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IView myListView = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Views.Add(new ViewOptions()
                                {
                                    Title = "My View",
                                    RowLimit = 10,
                                    SetAsDefaultView = true,
                                    ViewFields = 
                                       [ "DocIcon", "LinkFilenameNoMenu", "Modified" ]
                                });

        Console.WriteLine(myListView.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpPnPCoreSdk_UpdateViewForLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IView myView = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            Views.Where(vw => vw.Title == "My View").FirstOrDefault();

        if (myView != null)
        {
            myView.RowLimit = 20;
            myView.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsSpPnPCoreSdk_DeleteViewFromLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IView myView = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            Views.Where(vw => vw.Title == "My View").FirstOrDefault();

        if (myView != null)
        {
            myView.Delete();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 017

//gavdcodebegin 018
static void CsSpPnPCoreSdk_GetAllContentTypesLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentTypeCollection myListContentTypes = myContext.Web.Lists.GetByTitle(
                                                    "NewListPnPCoreSDK").ContentTypes;

        foreach (IContentType oneContentType in myListContentTypes)
        {
            Console.WriteLine(oneContentType.Name);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 018

//gavdcodebegin 019
static void CsSpPnPCoreSdk_EnableContentTypesListProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK",
                                                      lst => lst.ContentTypesEnabled);

        if (myList.ContentTypesEnabled == false)
        {
            myList.ContentTypesEnabled = true;
            myList.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 019

//gavdcodebegin 020
static void CsSpPnPCoreSdk_CreateContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType mySiteContentType = myContext.Web.ContentTypes.Add(
                    "0x010200A6A06C797CAAA84084CCA91D774D3B27", "MySiteContentType");

        IContentType myListContentType = myContext.Web.Lists
            .GetByTitle("NewListPnPCoreSDK").ContentTypes.AddAvailableContentType(
                                          "0x010200A6A06C797CAAA84084CCA91D774D3B27");

        Console.WriteLine(myListContentType.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 020

//gavdcodebegin 021
static void CsSpPnPCoreSdk_UpdateContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists
            .GetByTitle("NewListPnPCoreSDK").ContentTypes
            .Where(ct => ct.Name == "MySiteContentType").FirstOrDefault();

        if (myListContentType != null)
        {
            myListContentType.Name = "MyListContentType";
            myListContentType.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 021

//gavdcodebegin 022
static void CsSpPnPCoreSdk_AddFieldToContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists
            .GetByTitle("NewListPnPCoreSDK").ContentTypes
            .Where(ct => ct.Name == "MyListContentType").FirstOrDefault();

        if (myListContentType != null)
        {
            IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
            IField myField = myList.Fields.Where(
                           fld => fld.InternalName == "OneTextField").FirstOrDefault();

            myListContentType.FieldLinks.Add(myField, required: true);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 022

//gavdcodebegin 023
static void CsSpPnPCoreSdk_DeleteContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists
            .GetByTitle("NewListPnPCoreSDK").ContentTypes
            .Where(ct => ct.Name == "MyListContentType").FirstOrDefault();

        if (myListContentType != null)
        {
            myListContentType.Delete();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 023

//gavdcodebegin 024
static void CsSpPnPCoreSdk_BreakInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        if (myList != null)
        {
            myList.BreakRoleInheritance(false, true);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 024

//gavdcodebegin 025
static void CsSpPnPCoreSdk_HasInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.EnsureProperties(hur => hur.HasUniqueRoleAssignments);

        Console.WriteLine(myList.HasUniqueRoleAssignments.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 025

//gavdcodebegin 026
static void CsSpPnPCoreSdk_RestoreInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        if (myList != null)
        {
            myList.ResetRoleInheritance();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 026

//gavdcodebegin 027
static void CsSpPnPCoreSdk_GetAllSecurityRolesLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        foreach (IRoleAssignment oneRole in myList.RoleAssignments)
        {
            IRoleDefinitionCollection permLevels = myList.GetRoleDefinitions(
                                                                oneRole.PrincipalId);

            foreach (IRoleDefinition onePermLevel in permLevels)
            {
                Console.WriteLine(onePermLevel.Name);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 027

//gavdcodebegin 028
static void CsSpPnPCoreSdk_AddSecurityRoleToLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISharePointUser myUser = myContext.Web.GetCurrentUser();
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        myList.AddRoleDefinitions(myUser.Id, ["Read", "Edit"]);
        myList.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 028

//gavdcodebegin 029
static void CsSpPnPCoreSdk_DeleteSecurityRoleFromLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsPnpCoreSdk_GetContextWithAccPw(
        myTenantId, myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISharePointUser myUser = myContext.Web.GetCurrentUser();
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        myList.RemoveRoleDefinitions(myUser.Id, ["Read"]);
        myList.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 029


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 029 ***

//CsSpPnPCoreSdk_GetAllLists();
//CsSpPnPCoreSdk_GetOneList();
//CsSpPnPCoreSdk_GetOneListProperties();
//CsSpPnPCoreSdk_CreateList();
//CsSpPnPCoreSdk_UpdateList();
//CsSpPnPCoreSdk_RecycleList();
//CsSpPnPCoreSdk_DeleteList();
//CsSpPnPCoreSdk_GetAllFieldsInLists();
//CsSpPnPCoreSdk_GetPropertiesFieldsInLists();
//CsSpPnPCoreSdk_CreateFieldByXmlToLists();
//CsSpPnPCoreSdk_CreateFieldByApiToLists();
//CsSpPnPCoreSdk_UpdateFieldInLists();
//CsSpPnPCoreSdk_DeleteFieldFromLists();
//CsSpPnPCoreSdk_GetAllViewsInLists();
//CsSpPnPCoreSdk_CreateViewForLists();
//CsSpPnPCoreSdk_UpdateViewForLists();
//CsSpPnPCoreSdk_DeleteViewFromLists();
//CsSpPnPCoreSdk_GetAllContentTypesLists();
//CsSpPnPCoreSdk_EnableContentTypesListProperty();
//CsSpPnPCoreSdk_CreateContentTypeList();
//CsSpPnPCoreSdk_UpdateContentTypeList();
//CsSpPnPCoreSdk_AddFieldToContentTypeList();
//CsSpPnPCoreSdk_DeleteContentTypeList();
//CsSpPnPCoreSdk_BreakInheritanceList();
//CsSpPnPCoreSdk_HasInheritanceList();
//CsSpPnPCoreSdk_RestoreInheritanceList();
//CsSpPnPCoreSdk_GetAllSecurityRolesLists();
//CsSpPnPCoreSdk_AddSecurityRoleToLists();
//CsSpPnPCoreSdk_DeleteSecurityRoleFromLists();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
