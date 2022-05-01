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
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

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


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static void PnPCoreSdkGetAdminUrls()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        Uri myPortalUrl = myContext.GetSharePointAdmin().GetTenantPortalUri();
        Uri myAdminCenterUrl = myContext.GetSharePointAdmin().GetTenantAdminCenterUri();
        Uri myHostUrl = myContext.GetSharePointAdmin().GetTenantMySiteHostUri();

        Console.WriteLine(myPortalUrl + " - " + myAdminCenterUrl + " - " + myHostUrl);
    }
}
//gavdcodeend 01

//gavdcodebegin 02
static void PnPCoreSdkGetTenantProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
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
}
//gavdcodeend 02

//gavdcodebegin 03
static void PnPCoreSdkUpdateTenantProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITenantProperties myTenantProps = myContext.GetSharePointAdmin()
                                                                   .GetTenantProperties();

        if (myTenantProps.HideSyncButtonOnDocLib == false)
        {
            myTenantProps.HideSyncButtonOnDocLib = true;
            myTenantProps.Update();
        }
    }
}
//gavdcodeend 03

//gavdcodebegin 04
static void PnPCoreSdkGetTenantUsers()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        List<ISharePointUser> myTenantAdmins =
                                        myContext.GetSharePointAdmin().GetTenantAdmins();

        foreach (ISharePointUser myAdmin in myTenantAdmins)
        {
            Console.WriteLine(myAdmin.LoginName);
        }
    }
}
//gavdcodeend 04

//gavdcodebegin 05
static void PnPCoreSdkUserIsTenantAdmin()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        bool myUserIsAdmin = myContext.GetSharePointAdmin().IsCurrentUserTenantAdmin();

        Console.WriteLine(myUserName + " is Admin = " + myUserIsAdmin.ToString());
    }
}
//gavdcodeend 05

//gavdcodebegin 06
static void PnPCoreSdkHasTenantAppCatalog()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        bool myTenantHasAppCat = myContext.GetTenantAppManager().EnsureTenantAppCatalog();

        Console.WriteLine("Tenant has AppCatalog = " + myTenantHasAppCat.ToString());
    }
}
//gavdcodeend 06

//gavdcodebegin 07
static void PnPCoreSdkAppCatalogUrl()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        Uri myTenantAppCatUrl = myContext.GetTenantAppManager().GetTenantAppCatalogUri();

        Console.WriteLine("Tenant AppCatalog URL = " + myTenantAppCatUrl.ToString());
    }
}
//gavdcodeend 07

//gavdcodebegin 08
static void PnPCoreSdkGetAppCatalogs()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IList<IAppCatalogSite> myTenantAppCatUrl = 
                           myContext.GetTenantAppManager().GetSiteCollectionAppCatalogs();

        foreach (IAppCatalogSite myAppCat in myTenantAppCatUrl)
        {
            Console.WriteLine(myAppCat.AbsoluteUrl);
        }
    }
}
//gavdcodeend 08

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//PnPCoreSdkGetAdminUrls();
//PnPCoreSdkGetTenantProperties();
//PnPCoreSdkUpdateTenantProperty();
//PnPCoreSdkGetTenantUsers();
//PnPCoreSdkUserIsTenantAdmin();
//PnPCoreSdkHasTenantAppCatalog();
//PnPCoreSdkAppCatalogUrl();
PnPCoreSdkGetAppCatalogs();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
