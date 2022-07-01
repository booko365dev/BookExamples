using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth;
using PnP.Core.Model.SharePoint;
using PnP.Core.Model.Security;
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
static void SpCsPnPCoreSdk_GetAllLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IListCollection myLists = myContext.Web.Lists;

        foreach (IList oneList in myLists)
        {
            Console.WriteLine(oneList.Title);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsPnPCoreSdk_GetOneList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myOneList = myContext.Web.Lists.GetByTitle("Documents");
        //IList myOneList = myContext.Web.Lists.Where(
        //                              lst => lst.Title == "Documents").FirstOrDefault();
        //IList myOneList = myContext.Web.Lists.GetById(
        //                              new Guid("32243fc3-33dc-4b56-bec5-3166206c26ad"));
        //IList myOneList = myContext.Web.Lists.GetByServerRelativeUrl(
        //                              $"{myContext.Uri.PathAndQuery}/Shared Documents");

        Console.WriteLine(myOneList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsPnPCoreSdk_GetOneListProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("Documents", 
                                                      lst => lst.Id, 
                                                      lst => lst.TemplateType);

        Console.WriteLine(myList.Id + " - " + myList.TemplateType);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsPnPCoreSdk_CreateList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.Add(
                                    "NewListPnPCoreSDK", ListTemplateType.GenericList);

        Console.WriteLine(myList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsPnPCoreSdk_UpdateList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Description = "New Description for List";
        myList.Update();

        Console.WriteLine(myList.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsPnPCoreSdk_RecycleList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Recycle();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsPnPCoreSdk_DeleteList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 07

//gavdcodebegin 08
static void SpCsPnPCoreSdk_GetAllFieldsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 08

//gavdcodebegin 09
static void SpCsPnPCoreSdk_GetPropertiesFieldsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsPnPCoreSdk_CreateFieldByXmlToLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myFieldXml = 
            @"<Field Type=""Text"" Name=""myTextField"" DisplayName=""My Text Field""/>";

        IField myListField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Fields.AddFieldAsXml(myFieldXml, true);

        Console.WriteLine(myListField.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsPnPCoreSdk_CreateFieldByApiToLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IField myListField = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Fields.AddMultilineText("My MultilineField", 
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
//gavdcodeend 11

//gavdcodebegin 12
static void SpCsPnPCoreSdk_UpdateFieldInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 12

//gavdcodebegin 13
static void SpCsPnPCoreSdk_DeleteFieldFromLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsPnPCoreSdk_GetAllViewsInLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsPnPCoreSdk_CreateViewForLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IView myListView = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
                                Views.Add(new ViewOptions()
                                                    {
                                    Title = "My View",
                                    RowLimit = 10,
                                    SetAsDefaultView = true,
                                    ViewFields = new string[] { 
                                        "DocIcon", "LinkFilenameNoMenu", "Modified" }
                                });

        Console.WriteLine(myListView.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 15

//gavdcodebegin 16
static void SpCsPnPCoreSdk_UpdateViewForLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsPnPCoreSdk_DeleteViewFromLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsPnPCoreSdk_GetAllContentTypesLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsPnPCoreSdk_EnableContentTypesListProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK",
                                                      lst => lst.ContentTypesEnabled);

        if(myList.ContentTypesEnabled == false)
        {
            myList.ContentTypesEnabled = true;
            myList.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsPnPCoreSdk_CreateContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 20

//gavdcodebegin 21
static void SpCsPnPCoreSdk_UpdateContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            ContentTypes.Where(ct => ct.Name == "MySiteContentType").FirstOrDefault();

        if (myListContentType != null)
        {
            myListContentType.Name = "MyListContentType";
            myListContentType.Update();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsPnPCoreSdk_AddFieldToContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            ContentTypes.Where(ct => ct.Name == "MyListContentType").FirstOrDefault();

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
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsPnPCoreSdk_DeleteContentTypeList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IContentType myListContentType = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK").
            ContentTypes.Where(ct => ct.Name == "MyListContentType").FirstOrDefault();

        if (myListContentType != null)
        {
            myListContentType.Delete();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsPnPCoreSdk_BreakeInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        if (myList != null)
        {
            myList.BreakRoleInheritance(false, true);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 24

//gavdcodebegin 25
static async void SpCsPnPCoreSdk_HasInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");
        myList.EnsureProperties(hur => hur.HasUniqueRoleAssignments);

        Console.WriteLine(myList.HasUniqueRoleAssignments.ToString());
    }

    Console.WriteLine("Done");
}
//gavdcodeend 25

//gavdcodebegin 26
static void SpCsPnPCoreSdk_RestoreInheritanceList()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        if (myList != null)
        {
            myList.ResetRoleInheritance();
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 26

//gavdcodebegin 27
static void SpCsPnPCoreSdk_GetAllSecurityRolesLists()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        foreach (IRoleAssignment oneRole in myList.RoleAssignments)
        {
            IRoleDefinitionCollection permLevels = myList.GetRoleDefinitions(
                                                                oneRole.PrincipalId);

            foreach(IRoleDefinition onePermLevel in permLevels)
            {
                Console.WriteLine(onePermLevel.Name);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 27

//gavdcodebegin 28
static void SpCsPnPCoreSdk_AddSecurityRoleToLists()
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
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        myList.AddRoleDefinitions(myUser.Id, new string[] { "Read", "Edit" });
        myList.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 28

//gavdcodebegin 29
static void SpCsPnPCoreSdk_DeleteSecurityRoleFromLists()
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
        IList myList = myContext.Web.Lists.GetByTitle("NewListPnPCoreSDK");

        myList.RemoveRoleDefinitions(myUser.Id, new string[] { "Read" });
        myList.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 29


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsPnPCoreSdk_GetAllLists();
//SpCsPnPCoreSdk_GetOneList();
//SpCsPnPCoreSdk_GetOneListProperties();
//SpCsPnPCoreSdk_CreateList();
//SpCsPnPCoreSdk_UpdateList();
//SpCsPnPCoreSdk_RecycleList();
//SpCsPnPCoreSdk_DeleteList();
//SpCsPnPCoreSdk_GetAllFieldsInLists();
//SpCsPnPCoreSdk_GetPropertiesFieldsInLists();
//SpCsPnPCoreSdk_CreateFieldByXmlToLists();
//SpCsPnPCoreSdk_CreateFieldByApiToLists();
//SpCsPnPCoreSdk_UpdateFieldInLists();
//SpCsPnPCoreSdk_DeleteFieldFromLists();
//SpCsPnPCoreSdk_GetAllViewsInLists();
//SpCsPnPCoreSdk_CreateViewForLists();
//SpCsPnPCoreSdk_UpdateViewForLists();
//SpCsPnPCoreSdk_DeleteViewFromLists();
//SpCsPnPCoreSdk_GetAllContentTypesLists();
//SpCsPnPCoreSdk_EnableContentTypesListProperty();
//SpCsPnPCoreSdk_CreateContentTypeList();
//SpCsPnPCoreSdk_UpdateContentTypeList();
//SpCsPnPCoreSdk_AddFieldToContentTypeList();
//SpCsPnPCoreSdk_DeleteContentTypeList();
//SpCsPnPCoreSdk_BreakeInheritanceList();
//SpCsPnPCoreSdk_HasInheritanceList();
//SpCsPnPCoreSdk_RestoreInheritanceList();
//SpCsPnPCoreSdk_GetAllSecurityRolesLists();
//SpCsPnPCoreSdk_AddSecurityRoleToLists();
//SpCsPnPCoreSdk_DeleteSecurityRoleFromLists();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable