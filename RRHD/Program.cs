using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;

namespace RRHD
{
    class Program
    {
        static void Main(string[] args)
        {
            //SpCsPnPFrameworkExampleDel();     //==> PnP Framework Delegate permissions
            //SpCsPnPFrameworkExampleApp();     //==> PnP Framework Application permissions
            //SpCsPnPFrameworkExampleManag();   //==> PnP Framework Management Shell
            //SpCsPnPFrameworkExampleSecret();    //==> PnP Framework using Secret

            //LoginPnPFramework_UrlAppIdAppSecret();

            Console.ReadLine();
        }

        //-------------------------------------------------------------------------------

        //gavdcodebegin 05
        static void SpCsPnPFrameworkExampleDel()
        {
            using (ClientContext spPnpCtx = LoginPnPFramework_Delegate())
            {
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
        static void SpCsPnPFrameworkExampleApp()
        {
            using (ClientContext spPnpCtx = LoginPnPFramework_Application())
            {
                Web myWeb = spPnpCtx.Web;
                spPnpCtx.Load(myWeb, mw => mw.Id, mw => mw.Title);
                spPnpCtx.ExecuteQuery();

                Console.WriteLine(myWeb.Id + " - " + myWeb.Title);

                List myDocuments = myWeb.GetListByTitle("Documents", md => md.Id, md => md.Title);

                Console.WriteLine(myDocuments.Id + " - " + myDocuments.Title);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsPnPFrameworkExampleManag()
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
        static void SpCsPnPFrameworkExampleSecret()
        {
            using (ClientContext spPnpCtx = LoginPnPFramework_UrlAppIdAppSecret())
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

        //-------------------------------------------------------------------------------

        //gavdcodebegin 01
        static ClientContext LoginPnPFramework_Delegate()
        {
            SecureString mySecurePw = new SecureString();
            foreach (char oneChr in ConfigurationManager.AppSettings["spUserPw"])
            { mySecurePw.AppendChar(oneChr); }

            PnP.Framework.AuthenticationManager myAuthManager = new
                PnP.Framework.AuthenticationManager(
                                    ConfigurationManager.AppSettings["azAppIdDelegate"],
                                    ConfigurationManager.AppSettings["spUserName"],
                                    mySecurePw);

            ClientContext rtnContext = myAuthManager.GetContext(
                                    ConfigurationManager.AppSettings["spUrl"]);

            return rtnContext;
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static ClientContext LoginPnPFramework_Application()
        {
            PnP.Framework.AuthenticationManager myAuthManager = new
                PnP.Framework.AuthenticationManager(
                                    ConfigurationManager.AppSettings["azAppIdApplication"],
                                    @"[PathForTheCertificate]",
                                    "[PasswordForTheCertificate]",
                                    "[Domain].onmicrosoft.com");

            ClientContext rtnContext = myAuthManager.GetContext(
                                             ConfigurationManager.AppSettings["spUrl"]);

            return rtnContext;
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static ClientContext LoginPnPFramework_PnPManagementShell()
        {
            SecureString mySecurePw = new SecureString();
            foreach (char oneChr in ConfigurationManager.AppSettings["spUserPw"])
            { mySecurePw.AppendChar(oneChr); }

            PnP.Framework.AuthenticationManager myAuthManager = new
                PnP.Framework.AuthenticationManager(
                                    ConfigurationManager.AppSettings["spUserName"],
                                    mySecurePw);

            ClientContext rtnContext = myAuthManager.GetContext(
                                    ConfigurationManager.AppSettings["spUrl"]);

            return rtnContext;
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static ClientContext LoginPnPFramework_UrlAppIdAppSecret()
        {
            ClientContext rtnContext = new
                PnP.Framework.AuthenticationManager().GetACSAppOnlyContext(
                                ConfigurationManager.AppSettings["spUrl"],
                                ConfigurationManager.AppSettings["azAppIdApplication"],
                                ConfigurationManager.AppSettings["azAppSecret"]);

            return rtnContext;
        }
        //gavdcodeend 04
    }
}
