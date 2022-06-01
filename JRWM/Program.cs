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
static void PnPCoreSdkCreateCommunicationSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        CommunicationSiteOptions myCommSiteOptions = new CommunicationSiteOptions(
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
//gavdcodeend 01

//gavdcodebegin 02
static void PnPCoreSdkCreateTeamNoGroupSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        TeamSiteWithoutGroupOptions myTeamSiteOptions = new TeamSiteWithoutGroupOptions(
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
//gavdcodeend 02

//gavdcodebegin 03
static void PnPCoreSdkCreateTeamClassicSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ClassicSiteOptions myTeamSiteClassicOptions = new ClassicSiteOptions(
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
//gavdcodeend 03

//gavdcodebegin 04
static void PnPCoreSdkGetAllSiteColls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 04

//gavdcodebegin 05
static void PnPCoreSdkGetAllSiteCollsWithDetails()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 05

//gavdcodebegin 06
static void PnPCoreSdkGetOneSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionWithDetails mySiteColl = myContext.GetSiteCollectionManager().
                            GetSiteCollectionWithDetails(new Uri(mySiteCollUrl));

        Console.WriteLine(mySiteColl.Name + " - " + mySiteColl.TemplateName);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 06

//gavdcodebegin 07
static void PnPCoreSdkGetAllWebsInSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 07

//gavdcodebegin 08
static void PnPCoreSdkGetSiteCollProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionProperties mySiteCollProps = myContext.GetSiteCollectionManager().
                            GetSiteCollectionProperties(new Uri(mySiteCollUrl));

        var stringPropertyNamesAndValues = mySiteCollProps.GetType()
            .GetProperties()
            .Where(prop => prop.PropertyType == typeof(string) && prop.GetGetMethod() != null)
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
//gavdcodeend 08

//gavdcodebegin 09
static void PnPCoreSdkChangeOneSiteCollProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ISiteCollectionProperties mySiteCollProps = myContext.GetSiteCollectionManager().
                            GetSiteCollectionProperties(new Uri(mySiteCollUrl));

        mySiteCollProps.DisableFlows = FlowsPolicy.NotDisabled;
    }

    Console.WriteLine("Done");
}
//gavdcodeend 09

//gavdcodebegin 10
static void PnPCoreSdkConnectSiteCollToGroup()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ConnectSiteToGroupOptions myConnectGroup = new ConnectSiteToGroupOptions(
            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                    "/sites/NewTeamSiteCollFromPnPCoreSdk"),
            "GroupFor_NewTeamSiteCollFromPnPCoreSdk",
            "New Title For SiteColl NewTeamSiteCollFromPnPCoreSdk");
        myContext.GetSiteCollectionManager().ConnectSiteCollectionToGroup(myConnectGroup);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 10

//gavdcodebegin 11
static void PnPCoreSdkDeleteSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 11

//gavdcodebegin 11
static void PnPCoreSdkGetDeletedSiteColls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
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
//gavdcodeend 11

//gavdcodebegin 12
static void PnPCoreSdkGetRestoredSiteColl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        myContext.GetSiteCollectionManager().RestoreSiteCollection(
            new Uri(ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                    "/sites/NewTeamSiteCollFromPnPCoreSdk"));
    }

    Console.WriteLine("Done");
}
//gavdcodeend 12


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//PnPCoreSdkCreateCommunicationSiteColl();
//PnPCoreSdkCreateTeamNoGroupSiteColl();
//PnPCoreSdkCreateTeamClassicSiteColl();
//PnPCoreSdkGetAllSiteColls();
//PnPCoreSdkGetAllSiteCollsWithDetails();
//PnPCoreSdkGetOneSiteColl();
//PnPCoreSdkGetAllWebsInSiteColl();
//PnPCoreSdkGetSiteCollProperties();
//PnPCoreSdkChangeOneSiteCollProperty();
//PnPCoreSdkConnectSiteCollToGroup();
//PnPCoreSdkDeleteSiteColl();
//PnPCoreSdkGetDeletedSiteColls();
//PnPCoreSdkGetRestoredSiteColl();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable