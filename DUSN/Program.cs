using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Core.Auth;
using PnP.Core.Model.Security;
using PnP.Core.Services;
using System.Configuration;
using System.Security;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static PnPContext CsSpPnPCoreSdk_GetContextWithAccPw(string TenantId, string ClientId,
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


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpPnPCoreSdk_GetAdminUrls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);

    Uri myPortalUrl = myContext.GetSharePointAdmin().GetTenantPortalUri();
    Uri myAdminCenterUrl = myContext.GetSharePointAdmin().GetTenantAdminCenterUri();
    Uri myHostUrl = myContext.GetSharePointAdmin().GetTenantMySiteHostUri();

    Console.WriteLine(myPortalUrl + " - " + myAdminCenterUrl + " - " + myHostUrl);
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnPCoreSdk_GetTenantProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);

    ITenantProperties myTenantProps = myContext.GetSharePointAdmin()
                                                               .GetTenantProperties();

    IDictionary<String, Object> myTenantPropsDictionary = myTenantProps
                .GetType()
                .GetProperties()
                .Where(p => p.CanRead)
                .ToDictionary(p => p.Name, p => p.GetValue(myTenantProps, null));

    foreach (string oneProp in myTenantPropsDictionary.Keys)
    {
        Console.WriteLine(oneProp + " = " + myTenantPropsDictionary[oneProp]);
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnPCoreSdk_UpdateTenantProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);

    ITenantProperties myTenantProps = myContext.GetSharePointAdmin()
                                                               .GetTenantProperties();

    if (myTenantProps.HideSyncButtonOnDocLib == false)
    {
        myTenantProps.HideSyncButtonOnDocLib = true;
        myTenantProps.Update();
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnPCoreSdk_GetTenantUsers()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);

    List<ISharePointUser> myTenantAdmins =
                                    myContext.GetSharePointAdmin().GetTenantAdmins();

    foreach (ISharePointUser myAdmin in myTenantAdmins)
    {
        Console.WriteLine(myAdmin.LoginName);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnPCoreSdk_UserIsTenantAdmin()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);
 
    bool myUserIsAdmin = myContext.GetSharePointAdmin().IsCurrentUserTenantAdmin();

    Console.WriteLine(myUserName + " is Admin = " + myUserIsAdmin.ToString());
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnPCoreSdk_HasTenantAppCatalog()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);
 
    bool myTenantHasAppCat = myContext.GetTenantAppManager().EnsureTenantAppCatalog();

    Console.WriteLine("Tenant has AppCatalog = " + myTenantHasAppCat.ToString());
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnPCoreSdk_AppCatalogUrl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);

    Uri myTenantAppCatUrl = myContext.GetTenantAppManager().GetTenantAppCatalogUri();

    Console.WriteLine("Tenant AppCatalog URL = " + myTenantAppCatUrl.ToString());
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnPCoreSdk_GetAppCatalogs()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsSpPnPCoreSdk_GetContextWithAccPw(myTenantId, 
                    myClientId, myUserName, myUserPw, mySiteCollUrl, LogLevel.None);
 
    IList<IAppCatalogSite> myTenantAppCatUrl =
                       myContext.GetTenantAppManager().GetSiteCollectionAppCatalogs();

    foreach (IAppCatalogSite myAppCat in myTenantAppCatUrl)
    {
        Console.WriteLine(myAppCat.AbsoluteUrl);
    }
}
//gavdcodeend 008

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 008 ***

//CsSpPnPCoreSdk_GetAdminUrls();
//CsSpPnPCoreSdk_GetTenantProperties();
//CsSpPnPCoreSdk_UpdateTenantProperty();
//CsSpPnPCoreSdk_GetTenantUsers();
//CsSpPnPCoreSdk_UserIsTenantAdmin();
//CsSpPnPCoreSdk_HasTenantAppCatalog();
//CsSpPnPCoreSdk_AppCatalogUrl();
//CsSpPnPCoreSdk_GetAppCatalogs();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
