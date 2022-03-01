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

//gavdcodebegin 01
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
//gavdcodeend 01

//gavdcodebegin 02
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
//gavdcodeend 02

//gavdcodebegin 03
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
//gavdcodeend 03

//gavdcodebegin 04
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
//gavdcodeend 04

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 05
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
//gavdcodeend 05

//gavdcodebegin 06
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
//gavdcodeend 06

//gavdcodebegin 07
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
//gavdcodeend 07

//gavdcodebegin 08
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
//gavdcodeend 08

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

