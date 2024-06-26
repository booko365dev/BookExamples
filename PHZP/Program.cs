﻿using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Sites;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace PHZP
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = CsSpPnpcore_Login();

            //CsSpPnpcore_SiteIsCommunication(spCtx);
            //CsSpPnpcore_CreateOneCommunicationSiteCollection();
            //CsSpPnpcore_FindWebTemplates(spCtx);
            //CsSpPnpcore_CreateOneWebInSiteCollection(spCtx);
            //CsSpPnpcore_GetWebsInSiteCollection(spCtx);
            //CsSpPnpcore_WebExists(spCtx);
            //CsSpPnpcore_ExportSearchSettings();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 001
        static void CsSpPnpcore_SiteIsCommunication(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            bool SiteIsCommnication = spCtx.Site.IsCommunicationSite();
            Console.WriteLine(SiteIsCommnication);
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        static void CsSpPnpcore_CreateOneCommunicationSiteCollection()//*** LEGACY CODE ***
        {
            string myBaseUrl = ConfigurationManager.AppSettings["spBaseUrl"];
            ClientContext spCtx = CsSpPnpcore_Login(myBaseUrl);

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
        //gavdcodeend 002

        //gavdcodebegin 003
        static void CsSpPnpcore_FindWebTemplates(ClientContext spCtx)  //*** LEGACY CODE ***
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
        //gavdcodeend 003

        //gavdcodebegin 004
        static void CsSpPnpcore_CreateOneWebInSiteCollection(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;

            Web myWeb = mySite.RootWeb.CreateWeb("NewWebSiteModernCsPnP",
                                                "NewWebSiteModernCsPnP",
                                                "NewWebSiteModernCsPnP Description",
                                                "STS#3", 1033, true, true);
        }
        //gavdcodeend 004

        //gavdcodebegin 005
        static void CsSpPnpcore_GetWebsInSiteCollection(
                                            ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;

            IEnumerable<string> myWebs = mySite.GetAllWebUrls();

            foreach (string oneWeb in myWebs)
            {
                Console.WriteLine(oneWeb);
            }
        }
        //gavdcodeend 005

        //gavdcodebegin 006
        static void CsSpPnpcore_WebExists(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Site mySite = spCtx.Site;
            spCtx.Load(mySite);
            spCtx.ExecuteQuery();

            string webFullUrl = spCtx.Site.Url + "/NewWebSiteModernCsPnP";
            bool webExists = spCtx.WebExistsFullUrl(webFullUrl);
            Console.WriteLine(webExists);
        }
        //gavdcodeend 006

        //gavdcodebegin 007
        static void CsSpPnpcore_ExportSearchSettings()  //*** LEGACY CODE ***
        {
            string fullWebUrl = ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewCommSiteCollectionCsPnP";
            ClientContext webCtx = CsSpPnpcore_Login(fullWebUrl);

            webCtx.ExportSearchSettings(@"C:\Temporary\search.xml",
              Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel.SPWeb);
        }
        //gavdcodeend 007

        //-------------------------------------------------------------------------------
        static ClientContext CsSpPnpcore_Login()  //*** LEGACY CODE ***
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

        static ClientContext CsSpPnpcore_Login(string SiteFullUrl)  //*** LEGACY CODE ***
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
