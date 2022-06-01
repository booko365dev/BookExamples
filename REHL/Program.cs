using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;
using PnP.Framework.Sites;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext LoginPnPFramework_WithAccPw()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithCertificate()
{
    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_PnPManagementShell()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query
    ClientContext rtnContext = new
        AuthenticationManager().GetACSAppOnlyContext(
                        ConfigurationManager.AppSettings["SiteCollUrl"],
                        ConfigurationManager.AppSettings["ClientIdWithSecret"],
                        ConfigurationManager.AppSettings["ClientSecret"]);

    return rtnContext;
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static void SpCsPnpcoreSiteIsCommunication()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        bool SiteIsCommnication = spPnpCtx.Site.IsCommunicationSite();
        Console.WriteLine(SiteIsCommnication);
    }
}
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsPnpcoreCreateOneCommunicationSiteCollection()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        string myBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];

        CommunicationSiteCollectionCreationInformation mySiteCreationProps =
                                new CommunicationSiteCollectionCreationInformation
                                {
                                    Url = myBaseUrl + "/sites/NewCommSiteCollectionCsPnP",
                                    Title = "NewCommSiteCollectionCsPnP",
                                    Lcid = 1033,
                                    ShareByEmailEnabled = false,
                                    SiteDesign = CommunicationSiteDesign.Topic
                                };

        ClientContext spCommCtx = spPnpCtx.CreateSiteAsync(mySiteCreationProps).Result;
    }
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsPnpcoreFindWebTemplates()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Site mySite = spPnpCtx.Site;
        WebTemplateCollection myTemplates = mySite.GetWebTemplates(1033, 0);
        spPnpCtx.Load(myTemplates);
        spPnpCtx.ExecuteQuery();

        foreach (WebTemplate oneTemplate in myTemplates)
        {
            Console.WriteLine(oneTemplate.Name + " - " + oneTemplate.Title);
        }
    }
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsPnpcoreCreateOneWebInSiteCollection()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Site mySite = spPnpCtx.Site;

        Web myWeb = mySite.RootWeb.CreateWeb("NewWebSiteModernCsPnP",
                                            "NewWebSiteModernCsPnP",
                                            "NewWebSiteModernCsPnP Description",
                                            "STS#3", 1033, true, true);
    }
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsPnpcoreGetWebsInSiteCollection()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Site mySite = spPnpCtx.Site;

        IEnumerable<string> myWebs = mySite.GetAllWebUrls();

        foreach (string oneWeb in myWebs)
        {
            Console.WriteLine(oneWeb);
        }
    }
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsPnpcoreWebExists()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Site mySite = spPnpCtx.Site;
        spPnpCtx.Load(mySite);
        spPnpCtx.ExecuteQuery();

        string webFullUrl = spPnpCtx.Site.Url + "/NewWebSiteModernCsPnP";
        bool webExists = spPnpCtx.WebExistsFullUrl(webFullUrl);
        Console.WriteLine(webExists);
    }
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsPnpcoreExportSearchSettings()
{
    string fullWebUrl = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                                    "/sites/NewCommSiteCollectionCsPnP";

    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext webContext = myAuthManager.GetContext(fullWebUrl);

    using (webContext)
    {
        webContext.ExportSearchSettings(@"C:\Temporary\search.xml",
               Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel.SPWeb);
    }
}
//gavdcodeend 07


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsPnpcoreSiteIsCommunication();
//SpCsPnpcoreCreateOneCommunicationSiteCollection();
//SpCsPnpcoreFindWebTemplates();
//SpCsPnpcoreCreateOneWebInSiteCollection();
//SpCsPnpcoreGetWebsInSiteCollection();
//SpCsPnpcoreWebExists();
//SpCsPnpcoreExportSearchSettings();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable

