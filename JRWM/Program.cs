using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Auth;
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

static PnPContext CsSpPnpCoreSdk_GetContextWithInteraction(string TenantId, 
                        string ClientId, string SiteCollUrl, LogLevel ShowLogs)
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

static PnPContext CsSpPnpCoreSdk_GetContextWithAccPw(string TenantId, string ClientId,
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

static PnPContext CsSpPnpCoreSdk_GetContextWithCertificate(string TenantId, 
    string ClientId, string CertificateThumbprint, string SiteCollUrl, LogLevel ShowLogs)
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
static void CsSpPnpCoreSdk_CreateCommunicationSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        CommunicationSiteOptions myCommSiteOptions = new(
                            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                            "/sites/NewCommSiteCollFromPnPCoreSdk"),
                            "NewCommunicationSiteCollPnPCoreSdk")
        {
            Description = "Communication Site description",
            Language = Language.English
        };

        PnPContext ctxNewSiteColl = myContext.GetSiteCollectionManager()
                                            .CreateSiteCollection(myCommSiteOptions);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnpCoreSdk_CreateTeamNoGroupSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        TeamSiteWithoutGroupOptions myTeamSiteOptions = new(
                            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                            "/sites/NewTeamSiteCollFromPnPCoreSdk"),
                            "NewTeamSiteCollPnPCoreSdk")
        {
            Description = "Team Site description",
            Language = Language.English
        };

        PnPContext ctxNewSiteColl = myContext.GetSiteCollectionManager()
                                            .CreateSiteCollection(myTeamSiteOptions);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnpCoreSdk_CreateTeamClassicSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ClassicSiteOptions myTeamSiteClassicOptions = new(
                            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                            "/sites/NewClassicTeamSiteCollFromPnPCoreSdk"),
                            "NewClassicTeamSiteCollPnPCoreSdk",
                            "STS#3",
                            myUserName,
                            Language.English,
                            PnP.Core.Admin.Model.SharePoint.TimeZone
                                                    .UTCMINUS0500_BOGOTA_LIMA_QUITO);

        PnPContext ctxNewSiteColl = myContext.GetSiteCollectionManager()
                                        .CreateSiteCollection(myTeamSiteClassicOptions);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 003

//gavdcodebegin 014
static void CsSpPnpCoreSdk_CreateTeamSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId,
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        TeamSiteOptions myTeamSiteOptions = new(
                            "NewTeamSiteCollPnPCoreSdk",
                            "NewTeamSiteCollPnPCoreSdk")
        {
            Description = "Team Site description",
            Language = Language.English
        };

        PnPContext ctxNewSiteColl = myContext.GetSiteCollectionManager()
                                    .CreateSiteCollectionAsync(myTeamSiteOptions).Result;
    }

    Console.WriteLine("Done");
}
//gavdcodeend 014

//gavdcodebegin 004
static void CsSpPnpCoreSdk_GetAllSiteColls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        List<ISiteCollection> mySiteColls =
                            myContext.GetSiteCollectionManager().GetSiteCollections();

        foreach (ISiteCollection oneSiteColl in mySiteColls)
        {
            Console.WriteLine(oneSiteColl.Name);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnpCoreSdk_GetAllSiteCollsWithDetails()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        List<ISiteCollectionWithDetails> mySiteColls =
                    myContext.GetSiteCollectionManager().GetSiteCollectionsWithDetails();

        foreach (ISiteCollectionWithDetails oneSiteColl in mySiteColls)
        {
            Console.WriteLine(oneSiteColl.Name + " - " + oneSiteColl.TemplateName);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnpCoreSdk_GetOneSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionWithDetails mySiteColl = myContext.GetSiteCollectionManager().
                            GetSiteCollectionWithDetails(new Uri(mySiteCollUrl));

        Console.WriteLine(mySiteColl.Name + " - " + mySiteColl.TemplateName);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnpCoreSdk_GetAllWebsInSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        List<IWebWithDetails> mySiteCollWebs = myContext.GetSiteCollectionManager().
                            GetSiteCollectionWebsWithDetails(new Uri(mySiteCollUrl));

        foreach (IWebWithDetails oneWeb in mySiteCollWebs)
        {
            Console.WriteLine(oneWeb.Title + " - " + oneWeb.WebTemplate);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnpCoreSdk_GetSiteCollProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionProperties mySiteCollProps = myContext.GetSiteCollectionManager().
                            GetSiteCollectionProperties(new Uri(mySiteCollUrl));

        var stringPropertyNamesAndValues = mySiteCollProps.GetType()
            .GetProperties()
            .Where(prop => prop.PropertyType == typeof(string) && 
                                                            prop.GetGetMethod() != null)
            .Select(prop => new
            {
                Name = prop.Name,
                Value = prop.GetGetMethod().Invoke(mySiteCollProps, null)
            });
        foreach (var oneProp in stringPropertyNamesAndValues)
        {
            Console.WriteLine(oneProp.Name + " - " + oneProp.Value);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpPnpCoreSdk_ChangeOneSiteCollProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionProperties mySiteCollProps = myContext.GetSiteCollectionManager().
                            GetSiteCollectionProperties(new Uri(mySiteCollUrl));

        mySiteCollProps.DisableFlows = FlowsPolicy.NotDisabled;
    }

    Console.WriteLine("Done");
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpPnpCoreSdk_ConnectSiteCollToGroup()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ConnectSiteToGroupOptions myConnectGroup = new ConnectSiteToGroupOptions(
            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                    "/sites/NewTeamSiteCollFromPnPCoreSdk"),
            "GroupFor_NewTeamSiteCollFromPnPCoreSdk",
            "New Title For SiteColl NewTeamSiteCollFromPnPCoreSdk");
        myContext.GetSiteCollectionManager()
                                        .ConnectSiteCollectionToGroup(myConnectGroup);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpPnpCoreSdk_DeleteSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        myContext.GetSiteCollectionManager().DeleteSiteCollection(
            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                    "/sites/NewTeamSiteCollFromPnPCoreSdk"));
        //myContext.GetSiteCollectionManager().RecycleSiteCollection(
        //    new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
        //                            "/sites/NewTeamSiteCollFromPnPCoreSdk"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpPnpCoreSdk_GetDeletedSiteColls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        List<IRecycledSiteCollection> deletedSitColls = myContext
                            .GetSiteCollectionManager().GetRecycledSiteCollections();

        foreach (IRecycledSiteCollection oneDeleted in deletedSitColls)
        {
            Console.WriteLine(oneDeleted.Name);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpPnpCoreSdk_GetRestoredSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CsSpPnpCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        myContext.GetSiteCollectionManager().RestoreSiteCollection(
            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                    "/sites/NewTeamSiteCollFromPnPCoreSdk"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 013


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 014 ***

//CsSpPnpCoreSdk_CreateCommunicationSiteColl();
//CsSpPnpCoreSdk_CreateTeamNoGroupSiteColl();
//CsSpPnpCoreSdk_CreateTeamClassicSiteColl();
//CsSpPnpCoreSdk_CreateTeamSiteColl();
//CsSpPnpCoreSdk_GetAllSiteColls();
//CsSpPnpCoreSdk_GetAllSiteCollsWithDetails();
//CsSpPnpCoreSdk_GetOneSiteColl();
//CsSpPnpCoreSdk_GetAllWebsInSiteColl();
//CsSpPnpCoreSdk_GetSiteCollProperties();
//CsSpPnpCoreSdk_ChangeOneSiteCollProperty();
//CsSpPnpCoreSdk_ConnectSiteCollToGroup();
//CsSpPnpCoreSdk_DeleteSiteColl();
//CsSpPnpCoreSdk_GetDeletedSiteColls();
//CsSpPnpCoreSdk_GetRestoredSiteColl();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
