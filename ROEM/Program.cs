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

// Term Store

//gavdcodebegin 01
static void SpCsPnPCoreSdk_GetTermStore()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermStore myTermStore = myContext.TermStore.GetAsync().Result;

        Console.WriteLine(myTermStore.Id);

        List<string> storeLanguages = myTermStore.Languages;
        foreach(string oneLanguage in storeLanguages)
        {
            Console.WriteLine(oneLanguage);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsPnPCoreSdk_CreateTermGroup()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myNewGroup = myContext.TermStore.Groups.Add("CsPnpCoreSdkTermGroup", 
                                                               "Group description");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsPnPCoreSdk_FindTermGroups()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroupCollection myTermGroups = myContext.TermStore.Groups;
        foreach(ITermGroup oneTermGroup in myTermGroups)
        {
            Console.WriteLine(oneTermGroup.Id);
        }

        ITermGroup myTermGroup01 = myContext.TermStore.Groups.Where(
                                 p => p.Name == "CsPnpCoreSdkTermGroup").FirstOrDefault();
        Console.WriteLine(myTermGroup01.Id);

        ITermGroup myTermGroup02 = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        Console.WriteLine(myTermGroup02.Id);

        ITermGroup myTermGroup03 = myContext.TermStore.Groups.
                                          GetById("b577d217-e547-4b24-b428-1363dd5664c9");
        Console.WriteLine(myTermGroup03.Name);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsPnPCoreSdk_UpdateTermGroup()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        myTermGroup.Description = "Term Group description updated";
        myTermGroup.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsPnPCoreSdk_DeleteTermGroup()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        myTermGroup.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsPnPCoreSdk_CreateTermSet()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.Add("CsPnpCoreSdkTermSet", 
                                                  "Set description");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsPnPCoreSdk_FindTermSets()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");

        ITermSetCollection myTermSets = myTermGroup.Sets;
        foreach (ITermSet oneTermSet in myTermSets)
        {
            Console.WriteLine(oneTermSet.Id);
        }

        ITermSet myTermSet01 = myTermGroup.Sets.Where(
                    p => p.Id == "9406670e-e3df-4fc8-88fd-c1de6b6f8b4d").FirstOrDefault();
        Console.WriteLine(myTermSet01.Id);

        ITermSet myTermSet02 = myTermGroup.Sets.
                                          GetById("9406670e-e3df-4fc8-88fd-c1de6b6f8b4d");
        Console.WriteLine(myTermSet02.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 07

//gavdcodebegin 08
static void SpCsPnPCoreSdk_UpdateTermSet()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("9406670e-e3df-4fc8-88fd-c1de6b6f8b4d");
        myTermSet.Description = "Term Set description updated";
        myTermSet.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 08

//gavdcodebegin 09
static void SpCsPnPCoreSdk_DeleteTermSet()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("9406670e-e3df-4fc8-88fd-c1de6b6f8b4d");
        myTermSet.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsPnPCoreSdk_CreateTerm()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");
        ITerm myTerm = myTermSet.Terms.Add("CsPnpCoreSdkTerm", "Term description");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsPnPCoreSdk_FindTerms()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");

        ITermCollection myTerms = myTermSet.Terms;
        foreach (ITerm oneTerm in myTerms)
        {
            Console.WriteLine(oneTerm.Id);
        }

        ITerm myTerm01 = myTermSet.Terms.Where(
                    p => p.Id == "7fa28cea-39f9-4fa1-acab-8d6dbc89beb3").FirstOrDefault();
        Console.WriteLine(myTerm01.Id);

        ITerm myTerm02 = myTermSet.Terms.GetById("7fa28cea-39f9-4fa1-acab-8d6dbc89beb3");
        Console.WriteLine(myTerm02.Id);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 11

//gavdcodebegin 12
static void SpCsPnPCoreSdk_CreateSubTerm()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");
        ITerm myTerm = myTermSet.Terms.GetById("7fa28cea-39f9-4fa1-acab-8d6dbc89beb3");
        ITerm mySubTerm = myTerm.Terms.Add("CsPnpCoreSdkSubTerm", "Sub Term description");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 12

//gavdcodebegin 13
static void SpCsPnPCoreSdk_FindSubTerms()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");
        ITerm myTerm = myTermSet.Terms.GetById("7fa28cea-39f9-4fa1-acab-8d6dbc89beb3");

        ITermCollection mySubTerms = myTerm.Terms;
        foreach (ITerm oneSubTerm in mySubTerms.AsRequested())
        {
            Console.WriteLine(oneSubTerm.Id);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsPnPCoreSdk_UpdateTerm()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");
        ITerm myTerm = myTermSet.Terms.GetById("7fa28cea-39f9-4fa1-acab-8d6dbc89beb3");

        myTerm.AddProperty("MyPropertyKey", "MyPropertyValue");
        myTerm.Update();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsPnPCoreSdk_DeleteTerm()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        ITermGroup myTermGroup = myContext.TermStore.Groups.
                                                       GetByName("CsPnpCoreSdkTermGroup");
        ITermSet myTermSet = myTermGroup.Sets.
                                          GetById("dd0faa5b-943d-4f40-9938-94a974625f54");
        ITerm myTerm = myTermSet.Terms.GetById("7fa28cea-39f9-4fa1-acab-8d6dbc89beb3");

        myTerm.Delete();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 15

//----------------------------------------------------------------------------------------
// Search

//gavdcodebegin 16
static void SpCsPnPCoreSdk_Search()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string strToSearch = "teams";
        SearchOptions mySearchOptions = new SearchOptions(strToSearch)
        {
            TrimDuplicates = false,
            SelectProperties = new List<string>() { "Path", "Url", "Title", "ListId" }
        };

        ISearchResult allSrchResult = myContext.Web.Search(mySearchOptions);

        foreach (Dictionary<string, object> oneSrchResult in allSrchResult.Rows)
        {
            if (oneSrchResult["Title"].ToString().ToLower().Contains(strToSearch))
            {
                Console.WriteLine(oneSrchResult["Title"] + " - " + oneSrchResult["Url"]);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsPnPCoreSdk_SearchSorting()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string strToSearch = "teams";
        SearchOptions mySearchOptions = new SearchOptions(strToSearch)
        {
            TrimDuplicates = false,
            SelectProperties = new List<string>() { "Path", "Url", "Title", "ListId" },
            SortProperties = new List<SortOption>()
                {
                    new SortOption("Created", SortDirection.Descending),
                    new SortOption("ModifiedBy", SortDirection.Ascending)
                }
        };

        ISearchResult allSrchResult = myContext.Web.Search(mySearchOptions);

        foreach (Dictionary<string, object> oneSrchResult in allSrchResult.Rows)
        {
            if (oneSrchResult["Title"].ToString().ToLower().Contains(strToSearch))
            {
                Console.WriteLine(oneSrchResult["Title"] + " - " + oneSrchResult["Url"]);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsPnPCoreSdk_SearchRefiners()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string strToSearch = "teams";
        SearchOptions mySearchOptions = new SearchOptions(strToSearch)
        {
            TrimDuplicates = false,
            SelectProperties = new List<string>() { "Path", "Url", "Title", "ListId" },
            SortProperties = new List<SortOption>()
                {
                    new SortOption("Created", SortDirection.Descending)
                },
            RefineProperties = new List<string>() { "ModifiedBy" }
        };

        ISearchResult allSrchResult = myContext.Web.Search(mySearchOptions);

        //foreach (Dictionary<string, object> oneSrchResult in allSrchResult.Rows)
        //{
        //    if (oneSrchResult["Title"].ToString().ToLower().Contains(strToSearch))
        //    {
        //        Console.WriteLine(oneSrchResult["Title"] + " - " + oneSrchResult["Url"]);
        //    }
        //}

        foreach (var oneRefiner in allSrchResult.Refinements)
        {
            foreach (var oneRefinertResult in oneRefiner.Value)
            {
                Console.WriteLine(oneRefinertResult.Value + " - " + 
                                                                oneRefinertResult.Count);
            }
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsPnPCoreSdk_SearchExportConfig()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string xmlWebSearchConfig = myContext.Web.GetSearchConfigurationXml();

        string xmlSiteSearchConfig = myContext.Site.GetSearchConfigurationXml();
    }

    Console.WriteLine("Done");
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsPnPCoreSdk_SearchImportConfig()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string xmlWebSearchConfig = "Xml string with the Web Configuration";
        myContext.Web.SetSearchConfigurationXml(xmlWebSearchConfig);

        string xmlSiteSearchConfig = "Xml string with the SiteColl Configuration"; 
        myContext.Site.SetSearchConfigurationXml(xmlSiteSearchConfig);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 20

//----------------------------------------------------------------------------------------
// User Profile

//gavdcodebegin 21
static void SpCsPnPCoreSdk_GetMyProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IPersonProperties myProperties = myContext.Social.UserProfile.GetMyProperties();

        Dictionary<string, object> myUserProps = myProperties.UserProfileProperties;

        foreach (var oneUserProp in myUserProps)
        {
            Console.WriteLine(oneUserProp.Key + " - " + oneUserProp.Value);
        }
    }

    Console.WriteLine("Done");
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsPnPCoreSdk_GetSomeMyProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        IPersonProperties myProperties = myContext.Social.UserProfile
                                .GetMyProperties(p => p.DisplayName, p => p.AccountName);

        Console.WriteLine(myProperties.AccountName);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsPnPCoreSdk_GetOneUserProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myUser = "i:0#.f|membership|" + myUserName;
        IPersonProperties myProperties = myContext.Social.UserProfile
                                                            .GetPropertiesFor(myUser);

        Console.WriteLine(myProperties.Email);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsPnPCoreSdk_GetOnePropertyOneUserProperties()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myUser = "i:0#.f|membership|" + myUserName;
        string myProperty = myContext.Social.UserProfile.GetPropertyFor(myUser, "Email");

        Console.WriteLine(myProperty);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 24

//gavdcodebegin 25
static void SpCsPnPCoreSdk_ModifyOneSingleProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myUser = "i:0#.f|membership|" + myUserName;
        myContext.Social.UserProfile.SetSingleValueProfileProperty
                                            (myUser, "AboutMe", "Modified by PnPCore");
    }

    Console.WriteLine("Done");
}
//gavdcodeend 25

//gavdcodebegin 26
static void SpCsPnPCoreSdk_ModifyOneMultivalueProperty()
{
    string myTenantId = ConfigurationManager.AppSettings["TenantName"];
    string myClientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string mySiteCollUrl = ConfigurationManager.AppSettings["SiteCollUrl"];
    string myUserName = ConfigurationManager.AppSettings["UserName"];
    string myUserPw = ConfigurationManager.AppSettings["UserPw"];

    using (PnPContext myContext = CreateContextWithAccPw(myTenantId, myClientId,
                                    myUserName, myUserPw, mySiteCollUrl, LogLevel.None))
    {
        string myUser = "i:0#.f|membership|" + myUserName;
        var userSkills = new List<string>() { "SharePoint", "PowerShell" };
        myContext.Social.UserProfile.SetMultiValuedProfileProperty
                                                    (myUser, "SPS-Skills", userSkills);
    }

    Console.WriteLine("Done");
}
//gavdcodeend 26


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// CSOM Term Store
//SpCsPnPCoreSdk_GetTermStore();
//SpCsPnPCoreSdk_CreateTermGroup();
//SpCsPnPCoreSdk_FindTermGroups();
//SpCsPnPCoreSdk_UpdateTermGroup();
//SpCsPnPCoreSdk_DeleteTermGroup();
//SpCsPnPCoreSdk_CreateTermSet();
//SpCsPnPCoreSdk_FindTermSets();
//SpCsPnPCoreSdk_UpdateTermSet();
//SpCsPnPCoreSdk_DeleteTermSet();
//SpCsPnPCoreSdk_CreateTerm();
//SpCsPnPCoreSdk_FindTerms();
//SpCsPnPCoreSdk_CreateSubTerm();
//SpCsPnPCoreSdk_FindSubTerms();
//SpCsPnPCoreSdk_UpdateTerm();
//SpCsPnPCoreSdk_DeleteTerm();

// Search
//SpCsPnPCoreSdk_Search();
//SpCsPnPCoreSdk_SearchSorting();
//SpCsPnPCoreSdk_SearchRefiners();
//SpCsPnPCoreSdk_SearchExportConfig();
//SpCsPnPCoreSdk_SearchImportConfig();

// User Profile
//SpCsPnPCoreSdk_GetMyProperties();
//SpCsPnPCoreSdk_GetSomeMyProperties();
//SpCsPnPCoreSdk_GetOneUserProperties();
//SpCsPnPCoreSdk_GetOnePropertyOneUserProperties();
//SpCsPnPCoreSdk_ModifyOneSingleProperty();
//SpCsPnPCoreSdk_ModifyOneMultivalueProperty();

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

#nullable enable