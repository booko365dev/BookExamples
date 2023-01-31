using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
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
//gavdcodeend 001

//gavdcodebegin 002
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
//gavdcodeend 002

//gavdcodebegin 003
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
//gavdcodeend 003

//gavdcodebegin 004
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
//gavdcodeend 004

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 005
static void SpCsPnPFrameworkExampleWithAccPw()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        Web myWeb = spPnpCtx.Web;
        spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
        spPnpCtx.ExecuteQuery();

        Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

        List myDocuments = myWeb.GetListByTitle("Documents", md => md.Id, md => md.Title);

        Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void SpCsPnPFrameworkExampleWithCertificate()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithCertificate())
    {
        // Requires Application permissions for SharePoint - Sites.FullControl.All

        Site mySite = spPnpCtx.Site;
        spPnpCtx.Load(mySite, ms => ms.Id, ms => ms.RootWeb.Title);
        spPnpCtx.ExecuteQuery();

        Console.WriteLine(mySite.Id + " - " + mySite.RootWeb.Title);

        List myDocuments = mySite.RootWeb.GetListByTitle("Documents", md => md.Id, md => md.Title);

        Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void SpCsPnPFrameworkExampleWithManagementShell()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_PnPManagementShell())
    {
        Web myWeb = spPnpCtx.Web;
        spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
        spPnpCtx.ExecuteQuery();

        Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

        List myDocuments = myWeb.GetListByTitle("Documents", md => md.Id, md => md.Title);

        Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
    }
}
//gavdcodeend 007

//gavdcodebegin 008
static void SpCsPnPFrameworkExampleWithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query
    using (ClientContext spPnpCtx = LoginPnPFramework_WithSecret())
    {
        Web myWeb = spPnpCtx.Web;
        spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
        spPnpCtx.ExecuteQuery();

        Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

        List myDocuments = myWeb.GetListByTitle("Documents", md => md.Id, md => md.Title);

        Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
    }
}
//gavdcodeend 008

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------
//SpCsPnPFrameworkExampleWithAccPw();             //==> PnP Framework Delegate permissions
//SpCsPnPFrameworkExampleWithCertificate();       //==> PnP Framework Application permissions
//SpCsPnPFrameworkExampleWithManagementShell();   //==> PnP Framework Management Shell
//SpCsPnPFrameworkExampleWithSecret();            //==> PnP Framework using Secret

//LoginPnPFramework_UrlAppIdAppSecret();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable

