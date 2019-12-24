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

        static void SpCsCsomRemoveSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RemoveSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }

        static void SpCsCsomRestoreSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RestoreDeletedSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }

        static void SpCsCsomRemoveDeletedSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RemoveDeletedSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsCsom01");

            spAdminCtx.ExecuteQuery();
        }

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

        static void SpCsCsomRegisterAsHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.RegisterHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.ExecuteQuery();
        }

        static void SpCsCsomUnregisterAsHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.UnregisterHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");

            spAdminCtx.ExecuteQuery();
        }

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

        static void SpCsCsomremoveSiteFromHubSiteCollection(ClientContext spAdminCtx)
        {
            Tenant myTenant = new Tenant(spAdminCtx);
            myTenant.DisconnectSiteFromHubSite(
                ConfigurationManager.AppSettings["spBaseUrl"] +
                                             "/sites/NewSiteForHub");
            spAdminCtx.ExecuteQuery();
        }

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

        static void SpCsCsomDeleteOneWebInSiteCollection()
        {
            string myWebFullUrl = ConfigurationManager.AppSettings["spUrl"] +
                                                            "/NewWebSiteModernCsCsom";
            ClientContext spCtx = LoginCsom(myWebFullUrl);

            Web myWeb = spCtx.Web;
            myWeb.DeleteObject();
            spCtx.ExecuteQuery();
        }

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

