using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;
using PnP.Framework.Sites;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext LoginPnPFramework_WithAccPw()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnpFramework_LoginWithCertificate()
{
    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnPFramework_PnPManagementShell()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsSpPnPFramework_LoginWithSecret()  //*** LEGACY CODE ***
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

//gavdcodebegin 001
static void CsSpPnpFramework_SiteIsCommunication()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    bool SiteIsCommunication = spPnpCtx.Site.IsCommunicationSite();
    Console.WriteLine(SiteIsCommunication);
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnpFramework_CreateOneCommunicationSiteCollection()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    string myBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];

    CommunicationSiteCollectionCreationInformation mySiteCreationProps = new()
                    {
                        Url = myBaseUrl + "/sites/NewCommSiteCollectionCsPnP",
                        Title = "NewCommSiteCollectionCsPnP",
                        Lcid = 1033,
                        ShareByEmailEnabled = false,
                        SiteDesign = CommunicationSiteDesign.Topic
                    };

    ClientContext spCommCtx = spPnpCtx.CreateSiteAsync(mySiteCreationProps).Result;
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnpFramework_FindWebTemplates()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    Site mySite = spPnpCtx.Site;
    WebTemplateCollection myTemplates = mySite.GetWebTemplates(1033, 0);
    spPnpCtx.Load(myTemplates);
    spPnpCtx.ExecuteQuery();

    foreach (WebTemplate oneTemplate in myTemplates)
    {
        Console.WriteLine(oneTemplate.Name + " - " + oneTemplate.Title);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnpFramework_CreateOneWebInSiteCollection()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    Site mySite = spPnpCtx.Site;

    Web myWeb = mySite.RootWeb.CreateWeb("NewWebSiteModernCsPnP",
                                        "NewWebSiteModernCsPnP",
                                        "NewWebSiteModernCsPnP Description",
                                        "STS#3", 1033, true, true);
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnpFramework_GetWebsInSiteCollection()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    Site mySite = spPnpCtx.Site;

    IEnumerable<string> myWebs = mySite.GetAllWebUrls();

    foreach (string oneWeb in myWebs)
    {
        Console.WriteLine(oneWeb);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnpFramework_WebExists()
{
    using ClientContext spPnpCtx = LoginPnPFramework_WithAccPw();
    Site mySite = spPnpCtx.Site;
    spPnpCtx.Load(mySite);
    spPnpCtx.ExecuteQuery();

    string webFullUrl = spPnpCtx.Site.Url + "/NewWebSiteModernCsPnP";
    bool webExists = spPnpCtx.WebExistsFullUrl(webFullUrl);
    Console.WriteLine(webExists);
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnpFramework_ExportSearchSettings()
{
    string fullWebUrl = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                                    "/sites/NewCommSiteCollectionCsPnP";

    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
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
//gavdcodeend 007


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 007 ***

//CsSpPnpFramework_SiteIsCommunication();
//CsSpPnpFramework_CreateOneCommunicationSiteCollection();
//CsSpPnpFramework_FindWebTemplates();
//CsSpPnpFramework_CreateOneWebInSiteCollection();
//CsSpPnpFramework_GetWebsInSiteCollection();
//CsSpPnpFramework_WebExists();
//CsSpPnpFramework_ExportSearchSettings();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

