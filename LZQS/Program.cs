﻿using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Security;
using PnP.Framework;
using Microsoft.SharePoint.Client.Taxonomy;

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

//gavdcodebegin 001
static void SpCsPnPFramework_CreateTermGroup()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

        TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spPnpCtx);
        TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);

        TermGroup myTermGroup = myTermStore.CreateTermGroup("CsPnpFrameworkTermGroup");
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void SpCsPnPFramework_CreateTermGroupEnsure()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermGroup myTermGroup = spPnpCtx.Site.EnsureTermGroup
                                                           ("CsPnpFrameworkTermGroupEns");
        Console.WriteLine(myTermGroup.Id);
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void SpCsPnPFramework_FindTermGroup()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermGroup myTermGroup = spPnpCtx.Site.GetTermGroupByName
                                                           ("CsPnpFrameworkTermGroupEns");
        Console.WriteLine(myTermGroup.Id);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void SpCsPnPFramework_CreateTermSetEnsure()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermGroup myTermGroup = spPnpCtx.Site.EnsureTermGroup("CsPnpFrameworkTermGroup");
        TermSet myTermSet = myTermGroup.EnsureTermSet("CsPnpFrameworkTermSetEns");
        Console.WriteLine(myTermSet.Id);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsPnPFramework_FindTermSet()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermSetCollection myTermSet = spPnpCtx.Site.GetTermSetsByName(
                                                            "CsPnpFrameworkTermSetEns");
        Console.WriteLine(myTermSet[0].Id);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void SpCsPnPFramework_CreateTerm()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermGroup myTermGroup = spPnpCtx.Site.EnsureTermGroup("CsPnpFrameworkTermGroup");
        TermSet myTermSet = myTermGroup.EnsureTermSet("CsPnpFrameworkTermSetEns");
        Term myTerm = spPnpCtx.Site.AddTermToTermset(myTermSet.Id, "CsPnpFrameworkTerm");

        Console.WriteLine(myTerm.Id);
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void SpCsPnPFramework_FindTerm()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        TermSetCollection myTermSet = spPnpCtx.Site.GetTermSetsByName(
                                                            "CsPnpFrameworkTermSetEns");
        Term myTerm = spPnpCtx.Site.GetTermByName(myTermSet[0].Id, "CsPnpFrameworkTerm");

        Console.WriteLine(myTerm.Id);
    }
}
//gavdcodeend 007

//gavdcodebegin 008
static void SpCsPnPFramework_ExportTermStore()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        List<string> myTermStoreExport = spPnpCtx.Site.ExportAllTerms(true);
        foreach (string oneTerm in myTermStoreExport)
        {
            Console.WriteLine(oneTerm);
        }
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void SpCsPnPFramework_ImportTermStore()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        // Requires Delegated permissions for SharePoint - Sites.FullControl.All

        string[] myTerms = { "TermGroup01|TermSet01|Term01",
                             "TermGroup01|TermSet01|Term02" };

        spPnpCtx.Site.ImportTerms(myTerms, 1033);
    }
}
//gavdcodeend 009

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// PnP Framework Term Store
//SpCsPnPFramework_CreateTermGroup();
//SpCsPnPFramework_CreateTermGroupEnsure();
//SpCsPnPFramework_FindTermGroup();
//SpCsPnPFramework_CreateTermSetEnsure();
//SpCsPnPFramework_FindTermSet();
//SpCsPnPFramework_CreateTerm();
//SpCsPnPFramework_FindTerm();
//SpCsPnPFramework_ExportTermStore();
//SpCsPnPFramework_ImportTermStore();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable

