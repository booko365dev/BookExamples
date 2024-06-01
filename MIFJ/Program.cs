using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PnP.Core.Auth;
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

//gavdcodebegin 001
static PnPContext CsPnPCoreSdk_GetContextWithInteraction(string TenantId, 
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
//gavdcodeend 001

//gavdcodebegin 003
static PnPContext CsPnPCoreSdk_GetContextWithAccPw(string TenantId, string ClientId,
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
//gavdcodeend 003

//gavdcodebegin 005
static PnPContext CsPnPCoreSdk_GetContextWithCertificate(string TenantId, 
                string ClientId, string CertificateThumbprint, string SiteCollUrl, 
                LogLevel ShowLogs)
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
//gavdcodeend 005

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 002
static void CsPnPCoreSdk_GetWebWithInteraction()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];

    using PnPContext myContext = CsPnPCoreSdk_GetContextWithInteraction(myTenantId, myClientId,
                                                         mySiteCollUrl, LogLevel.None);
    myContext.Web.LoadAsync(p => p.Title).Wait();
    Console.WriteLine($"The title of the web is '" + myContext.Web.Title + "'");
}
//gavdcodeend 002

//gavdcodebegin 004
static void CsPnPCoreSdk_GetListsWithAccPw()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using PnPContext myContext = CsPnPCoreSdk_GetContextWithAccPw(myTenantId, myClientId,
                                   myUserName, myUserPw, mySiteCollUrl, LogLevel.Trace);
    myContext.Web.Lists.LoadAsync().Wait();
    foreach (IList oneList in myContext.Web.Lists)
    {
        Console.WriteLine("List - " + oneList.Title);
    }
}
//gavdcodeend 004

//gavdcodebegin 006
static void CsPnPCoreSdk_GetItemsWithCertificate()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myCertThumbprint = ConfigurationManager.AppSettings["CertificateThumbprint"];

    using PnPContext myContext = CsPnPCoreSdk_GetContextWithCertificate(myTenantId, myClientId,
                                      myCertThumbprint, mySiteCollUrl, LogLevel.Debug);
    myContext.Web.LoadAsync(p => p.Title).Wait();
    Console.WriteLine($"The title of the web is {myContext.Web.Title}");

    IList myList = myContext.Web.Lists.GetByTitle("Documents",
                            p => p.Title,
                            p => p.Items.QueryProperties(p => p.Title));

    foreach (IListItem oneItem in myList.Items)
    {
        Console.WriteLine("Item - " + oneItem.Title);
    }
}
//gavdcodeend 006

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 006 ***

//CsPnPCoreSdk_GetWebWithInteraction();
//CsPnPCoreSdk_GetListsWithAccPw();
//CsPnPCoreSdk_GetItemsWithCertificate();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
