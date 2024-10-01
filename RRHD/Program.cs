using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static ClientContext CsSpPnPFramework_GetContextWithAccPw()
{
    SecureString mySecurePw = new ();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new (
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}
//gavdcodeend 001

//gavdcodebegin 002
static ClientContext CsSpPnPFramework_GetContextWithCertificate()
{
    AuthenticationManager myAuthManager = new (
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            ConfigurationManager.AppSettings["CertificateFilePath"],
                            ConfigurationManager.AppSettings["CertificateFilePw"],
                            ConfigurationManager.AppSettings["TenantName"]);

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}
//gavdcodeend 002

//gavdcodebegin 003
static ClientContext CsSpPnPFramework_GetContextWithManagShell()  //*** LEGACY CODE ***
{
    SecureString mySecurePw = new ();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new (
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}
//gavdcodeend 003

//gavdcodebegin 004
static ClientContext CsSpPnPFramework_GetContextWithSecret()  //*** LEGACY CODE ***
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
static void CsSpPnPFramework_ExampleWithAccPw()
{
    // Requires Delegated permissions for SharePoint - Sites.FullControl.All

    using ClientContext spPnpCtx = CsSpPnPFramework_GetContextWithAccPw();

    Web myWeb = spPnpCtx.Web;
    spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
    spPnpCtx.ExecuteQuery();

    Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

    List myDocuments = myWeb.GetListByTitle("Documents",
                                        md => md.Id, md => md.Title);

    Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnPFramework_ExampleWithCertificate()
{
    // Requires Application permissions for SharePoint - Sites.FullControl.All

    using ClientContext spPnpCtx = CsSpPnPFramework_GetContextWithCertificate();

    Site mySite = spPnpCtx.Site;
    spPnpCtx.Load(mySite, ms => ms.Id, ms => ms.RootWeb.Title);
    spPnpCtx.ExecuteQuery();

    Console.WriteLine(mySite.Id + " - " + mySite.RootWeb.Title);

    List myDocuments = mySite.RootWeb.GetListByTitle("Documents",
                                            md => md.Id, md => md.Title);

    Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnPFramework_ExampleWithManagementShell()  //*** LEGACY CODE ***
{
    using ClientContext spPnpCtx = CsSpPnPFramework_GetContextWithManagShell();

    Web myWeb = spPnpCtx.Web;
    spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
    spPnpCtx.ExecuteQuery();

    Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

    List myDocuments = myWeb.GetListByTitle("Documents",
                                                md => md.Id, md => md.Title);

    Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnPFramework_ExampleWithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query

    using ClientContext spPnpCtx = CsSpPnPFramework_GetContextWithSecret();

    Web myWeb = spPnpCtx.Web;
    spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
    spPnpCtx.ExecuteQuery();

    Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

    List myDocuments = myWeb.GetListByTitle("Documents",
                                                md => md.Id, md => md.Title);

    Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
}
//gavdcodeend 008

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 008 ***

//CsSpPnPFramework_ExampleWithAccPw();             //==> PnP Framework Delegate permissions
//CsSpPnPFramework_ExampleWithCertificate();       //==> PnP Framework Application permissions
//CsSpPnPFramework_ExampleWithManagementShell();   //==> PnP Framework Management Shell  //*** LEGACY CODE ***
//CsSpPnPFramework_ExampleWithSecret();            //==> PnP Framework using Secret

//LoginPnPFramework_UrlAppIdAppSecret();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

