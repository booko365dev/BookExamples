using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Security;

namespace DWBW
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginCsom();
            ClientContext spAdminCtx = LoginAdminCsom();

            //SpCsCsomCreateOneSiteCollection(spAdminCtx);
            //SpCsCsomFindWebTemplates(spAdminCtx);
            //SpCsCsomReadAllSiteCollections(spAdminCtx);
            //SpCsCsomRemoveSiteCollection(spAdminCtx);
            //SpCsCsomRestoreSiteCollection(spAdminCtx);
            //SpCsCsomRemoveDeletedSiteCollection(spAdminCtx);
            //SpCsCsomCreateGroupForSite(spAdminCtx);
            //SpCsCsomSetAdministratorSiteCollection(spAdminCtx);
            //SpCsCsomRegisterAsHubSiteCollection(spAdminCtx);
            //SpCsCsomUnregisterAsHubSiteCollection(spAdminCtx);
            //SpCsCsomGetHubSiteCollectionProperties(spAdminCtx);
            //SpCsCsomUpdateHubSiteCollectionProperties(spAdminCtx);
            //SpCsCsomAddSiteToHubSiteCollection(spAdminCtx);
            //SpCsCsomremoveSiteFromHubSiteCollection(spAdminCtx);

            //SpCsCsomCreateOneWebInSiteCollection(spCtx);
            //SpCsCsomGetWebsInSiteCollection(spCtx);
            //SpCsCsomGetOneWebInSiteCollection();
            //SpCsCsomUpdateOneWebInSiteCollection();
            //SpCsCsomDeleteOneWebInSiteCollection();
            //SpCsCsomBreakSecurityInheritanceWeb();
            //SpCsCsomResetSecurityInheritanceWeb();
            //SpCsCsomAddUserToSecurityRoleInWeb();
            //SpCsCsomUpdateUserSecurityRoleInWeb();
            //SpCsCsomDeleteUserFromSecurityRoleInWeb();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsCsomCreateOneSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            string myUser = ConfigurationManager.AppSettings["spUserName"];
            SiteCreationProperties mySiteCreationProps = new SiteCreationProperties
            {
                Url = ConfigurationManager.AppSettings["spBaseUrl"] + 
                                                "/sites/NewSiteCollectionModernCsCsom01",
                Title = "NewSiteCollectionModernCsCsom01",
                Owner = ConfigurationManager.AppSettings["spUserName"],
                Template = "STS#3",
                StorageMaximumLevel = 100,
                UserCodeMaximumLevel = 50
            };

            SpoOperation myOps = myTenant.CreateSite(mySiteCreationProps);
            spAdminCtx.Load(myOps, ic => ic.IsComplete);
            spAdminCtx.ExecuteQuery();

            while (myOps.IsComplete == false)
            {
                System.Threading.Thread.Sleep(5000);
                myOps.RefreshLoad();
                spAdminCtx.ExecuteQuery();
            }
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsCsomFindWebTemplates(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            SPOTenantWebTemplateCollection myTemplates = 
                                    myTenant.GetSPOTenantWebTemplates(1033, 0);
            spAdminCtx.Load(myTemplates);
            spAdminCtx.ExecuteQuery();

            foreach (SPOTenantWebTemplate oneTemplate in myTemplates)
            {
                Console.WriteLine(oneTemplate.Name + " - " + oneTemplate.Title);
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsCsomReadAllSiteCollections(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.GetSiteProperties(0, true);

            SPOSitePropertiesEnumerable myProps = myTenant.GetSiteProperties(0, true);
            spAdminCtx.Load(myProps);
            spAdminCtx.ExecuteQuery();

            foreach (var oneSiteColl in myProps)
            {
                Console.WriteLine(oneSiteColl.Title + " - " + oneSiteColl.Url);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsCsomRemoveSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RemoveSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsCsomRestoreSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RestoreDeletedSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsCsomRemoveDeletedSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RemoveDeletedSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsCsomCreateGroupForSite(ClientContext spAdminCtx)
        {
            string[] myOwners = new string[] { "user@domain.onmicrosoft.com" };
            GroupCreationParams myGroupParams = new GroupCreationParams(spAdminCtx);
            myGroupParams.Owners = myOwners;
            //GroupCreationParams
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.CreateGroupForSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01",
                "GroupForNewSiteCollectionModernCsCsom01",
                "GroupForNewSiteCollAlias",
                true,
                myGroupParams);

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsCsomSetAdministratorSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.SetSiteAdmin(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01",
                "user@domain.onmicrosoft.com",
                true);

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsCsomRegisterAsHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RegisterHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsCsomUnregisterAsHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.UnregisterHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsCsomGetHubSiteCollectionProperties(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            HubSiteProperties myProps = myTenant.GetHubSitePropertiesByUrl(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.Load(myProps);
            spAdminCtx.ExecuteQuery();

            Console.WriteLine(myProps.Title);
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsCsomUpdateHubSiteCollectionProperties(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            HubSiteProperties myProps = myTenant.GetHubSitePropertiesByUrl(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.Load(myProps);
            spAdminCtx.ExecuteQuery();

            myProps.Title = myProps.Title + "_Updated";
            myProps.Update();

            spAdminCtx.Load(myProps);
            spAdminCtx.ExecuteQuery();

            Console.WriteLine(myProps.Title);
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsCsomAddSiteToHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.ConnectSiteToHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewSiteForHub",
            ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");
            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void SpCsCsomremoveSiteFromHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.DisconnectSiteFromHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewSiteForHub");
            spAdminCtx.ExecuteQuery();
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        static void SpCsCsomCreateOneWebInSiteCollection(ClientContext spCtx)
        {
            Site mySite = spCtx.Site;

            WebCreationInformation myWebCreationInfo = new WebCreationInformation
            {
                Url = "NewWebSiteModernCsCsom",
                Title = "NewWebSiteModernCsCsom",
                Description = "NewWebSiteModernCsCsom Description",
                UseSamePermissionsAsParentSite = true,
                WebTemplate = "STS#3",
                Language = 1033
            };

            Web myWeb = mySite.RootWeb.Webs.Add(myWebCreationInfo);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 15

        //gavdcodebegin 16
        static void SpCsCsomGetWebsInSiteCollection(ClientContext spCtx)
        {
            Site mySite = spCtx.Site;

            WebCollection myWebs = mySite.RootWeb.Webs;
            spCtx.Load(myWebs);
            spCtx.ExecuteQuery();

            foreach (Web oneWeb in myWebs)
            {
                Console.WriteLine(oneWeb.Title + " - " + oneWeb.Url + " - " + oneWeb.Id);
            }
        }
        //gavdcodeend 16

        //gavdcodebegin 17
        static void SpCsCsomGetOneWebInSiteCollection()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            spCtx.Load(myWeb);
            spCtx.ExecuteQuery();

            Console.WriteLine(myWeb.Title + " - " + myWeb.Url + " - " + myWeb.Id);
        }
        //gavdcodeend 17

        //gavdcodebegin 18
        static void SpCsCsomUpdateOneWebInSiteCollection()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            myWeb.Description = "NewWebSiteModernCsCsom Description Updated";
            myWeb.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        static void SpCsCsomDeleteOneWebInSiteCollection()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            myWeb.DeleteObject();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 19

        //gavdcodebegin 20
        static void SpCsCsomBreakSecurityInheritanceWeb()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            spCtx.Load(myWeb, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myWeb.HasUniqueRoleAssignments == false)
            {
                myWeb.BreakRoleInheritance(false, true);
            }
            myWeb.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 20

        //gavdcodebegin 21
        static void SpCsCsomResetSecurityInheritanceWeb()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            spCtx.Load(myWeb, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myWeb.HasUniqueRoleAssignments == true)
            {
                myWeb.ResetRoleInheritance();
            }
            myWeb.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 21

        //gavdcodebegin 22
        static void SpCsCsomAddUserToSecurityRoleInWeb()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
            myWeb.RoleAssignments.Add(myUser, roleDefinition);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 22

        //gavdcodebegin 23
        static void SpCsCsomUpdateUserSecurityRoleInWeb()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Administrator));

            RoleAssignment myRoleAssignment = myWeb.RoleAssignments.GetByPrincipal(
                                                                                myUser);
            myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

            myRoleAssignment.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 23

        //gavdcodebegin 24
        static void SpCsCsomDeleteUserFromSecurityRoleInWeb()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            myWeb.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

            spCtx.ExecuteQuery();
            spCtx.Dispose();
        }
        //gavdcodeend 24

        //-------------------------------------------------------------------------------
        static ClientContext LoginCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spUrl"]);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }

        static ClientContext LoginCsom(string WebFullUrl)
        {
            ClientContext rtnContext = new ClientContext(WebFullUrl);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }

        static ClientContext LoginAdminCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spAdminUrl"]);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }
    }
}
