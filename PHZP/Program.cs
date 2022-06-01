using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Sites;

namespace PHZP
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginPnPCore();

            //SpCsPnpcoreSiteIsCommunication(spCtx);
            //SpCsPnpcoreCreateOneCommunicationSiteCollection();
            //SpCsPnpcoreFindWebTemplates(spCtx);
            //SpCsPnpcoreCreateOneWebInSiteCollection(spCtx);
            //SpCsPnpcoreGetWebsInSiteCollection(spCtx);
            //SpCsPnpcoreWebExists(spCtx);
            //SpCsPnpcoreExportSearchSettings();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsPnpcoreSiteIsCommunication(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            bool SiteIsCommnication = spCtx.Site.IsCommunicationSite();
            Console.WriteLine(SiteIsCommnication);
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsPnpcoreCreateOneCommunicationSiteCollection()//*** LEGACY CODE ***
        {
            string myBaseUrl = ConfigurationManager.AppSettings["spBaseUrl"];
            ClientContext spCtx = LoginPnPCore(myBaseUrl);

            CommunicationSiteCollectionCreationInformation mySiteCreationProps = 
                                    new CommunicationSiteCollectionCreationInformation
            {
                Url = myBaseUrl + "/sites/NewCommSiteCollectionCsPnP",
                Title = "NewCommSiteCollectionCsPnP",
                Lcid = 1033,
                ShareByEmailEnabled = false,
                SiteDesign = CommunicationSiteDesign.Topic
            };

            ClientContext spCommCtx = spCtx.CreateSiteAsync(mySiteCreationProps).Result;
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsPnpcoreFindWebTemplates(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;
            WebTemplateCollection myTemplates = mySite.GetWebTemplates(1033, 0);
            spCtx.Load(myTemplates);
            spCtx.ExecuteQuery();

            foreach (WebTemplate oneTemplate in myTemplates)
            {
                Console.WriteLine(oneTemplate.Name + " - " + oneTemplate.Title);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsPnpcoreCreateOneWebInSiteCollection(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;

            Web myWeb = mySite.RootWeb.CreateWeb("NewWebSiteModernCsPnP", 
                                                "NewWebSiteModernCsPnP", 
                                                "NewWebSiteModernCsPnP Description", 
                                                "STS#3", 1033, true, true);
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsPnpcoreGetWebsInSiteCollection(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;

            IEnumerable<string> myWebs = mySite.GetAllWebUrls();

            foreach (string oneWeb in myWebs)
            {
                Console.WriteLine(oneWeb);
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsPnpcoreWebExists(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;
            spCtx.Load(mySite);
            spCtx.ExecuteQuery();

            string webFullUrl = spCtx.Site.Url + "/NewWebSiteModernCsPnP";
            bool webExists = spCtx.WebExistsFullUrl(webFullUrl);
            Console.WriteLine(webExists);
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsPnpcoreExportSearchSettings()  //*** LEGACY CODE ***
        {
            string fullWebUrl = ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewCommSiteCollectionCsPnP";
            ClientContext webCtx = LoginPnPCore(fullWebUrl);

            webCtx.ExportSearchSettings(@"C:\Temporary\search.xml", 
              Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel.SPWeb);
        }
        //gavdcodeend 07

        //-------------------------------------------------------------------------------
        static ClientContext LoginPnPCore()  //*** LEGACY CODE ***
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }

        static ClientContext LoginPnPCore(string SiteFullUrl)  //*** LEGACY CODE ***
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (SiteFullUrl,
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }
    }
}
