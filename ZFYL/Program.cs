using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace ZFYL
{
    class Program
    {
        static void Main(string[] args)
        {
            // CSOM Term Store
            //ClientContext spCtx = LoginCsom();
            //SpCsCsomFindTermStore(spCtx);
            //SpCsCsomCreateTermGroup(spCtx);
            //SpCsCsomFindTermGroups(spCtx);
            //SpCsCsomCreateTermSet(spCtx);
            //SpCsCsomFindTermSets(spCtx);
            //SpCsCsomCreateTerm(spCtx);
            //SpCsCsomFindTerms(spCtx);
            //SpCsCsomFindOneTerm(spCtx);
            //SpCsCsomUpdateOneTerm(spCtx);
            //SpCsCsomFindTermSetAndTermById(spCtx);
            //SpCsCsomDeleteOneTerm(spCtx);

            // PnPCore Term Store
            //ClientContext spCtxPnp = LoginPnPCore();
            //SpCsPnpcoreCreateTermGroup(spCtxPnp);
            //SpCsPnpcoreCreateTermGroupEnsure(spCtxPnp);
            //SpCsPnpcoreCreateTermSetEnsure(spCtxPnp);
            //SpCsPnpcoreCreateTerm(spCtxPnp);
            //SpCsPnpcoreFindTermGroup(spCtxPnp);
            //SpCsPnpcoreFindTermSet(spCtxPnp);
            //SpCsPnpcoreFindTerm(spCtxPnp);
            //SpCsPnpcoreExportTermStore(spCtxPnp);
            //SpCsPnpcoreImportTermStore(spCtxPnp);

            // CSOM Search
            //ClientContext spCtx = LoginCsom();
            //SpCsCsomGetResultsSearch(spCtx);

            // REST Search
            //Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            //string userName = ConfigurationManager.AppSettings["spUserName"];
            //string password = ConfigurationManager.AppSettings["spUserPw"];
            //SpCsRestResultsSearchGET(webUri, userName, password);
            //SpCsRestResultsSearchPOST(webUri, userName, password);

            // CSOM User Profile
            //ClientContext spCtx = LoginCsom();
            //SpCsCsomGetAllPropertiesUserProfile(spCtx);
            //SpCsCsomGetAllMyPropertiesUserProfile(spCtx);
            //SpCsCsomGetPropertiesUserProfile(spCtx);
            //SpCsCsomUpdateOnePropertyUserProfile(spCtx);
            //SpCsCsomUpdateOneMultPropertyUserProfile(spCtx);

            // REST User Profile
            //Uri webUri = new Uri(ConfigurationManager.AppSettings["spUrl"]);
            //string userName = ConfigurationManager.AppSettings["spUserName"];
            //string password = ConfigurationManager.AppSettings["spUserPw"];
            //SpCsRestGetAllPropertiesUserProfile(webUri, userName, password);
            //SpCsRestGetAllMyPropertiesUserProfile(webUri, userName, password);
            //SpCsRestGetPropertiesUserProfile(webUri, userName, password);

            // CSOM Site Scripts
            //ClientContext spCtx = LoginCsomAdminSite();
            //SpCsCsomGenerateWebSiteScript(spCtx);
            //SpCsCsomAddSiteScript(spCtx);
            //SpCsCsomGetAllSiteScripts(spCtx);  // Not working
            //SpCsCsomUpdateSiteScript(spCtx);   // Not working
            //SpCsCsomDeleteSiteScript(spCtx);

            // CSOM Site Designs
            //ClientContext spCtx = LoginCsomAdminSite();
            //SpCsCsomAddSiteDesign(spCtx);
            //SpCsCsomApplySiteDesign(spCtx);
            //SpCsCsomGetAllSiteDesigns(spCtx);
            //SpCsCsomUpdateSiteDesign(spCtx);
            //SpCsCsomGetTasksSiteDesign(spCtx);
            //SpCsCsomGetRunsSiteDesign(spCtx);
            //SpCsCsomGetRunStatusSiteDesign(spCtx);
            //SpCsCsomGrantRightsSiteDesign(spCtx);
            //SpCsCsomDeleteSiteDesign(spCtx);

            // PnPCore Provisioning
            //ClientContext spCtxPnp = LoginPnPCore();
            //SpCsPnpcoreGenerateSiteTemplateXml(spCtxPnp);
            //SpCsPnpcoreGenerateSiteListTemplate(spCtxPnp);
            //SpCsPnpcoreApplySiteTemplate(spCtxPnp);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsCsomFindTermStore(ClientContext spCtx)
        {
            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            spCtx.Load(myTaxSession, ts => ts.TermStores);
            spCtx.ExecuteQuery();

            foreach (TermStore oneTermStore in myTaxSession.TermStores)
            {
                Console.WriteLine(oneTermStore.Name);
            }
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsCsomCreateTermGroup(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);

            TermGroup myTermGroup = myTermStore.CreateGroup(
                                                    "CsCsomTermGroup", Guid.NewGuid());
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsCsomFindTermGroups(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            spCtx.Load(myTermStore, tStore => tStore.Name, tStore => tStore.Groups);
            spCtx.ExecuteQuery();

            foreach (TermGroup oneGroup in myTermStore.Groups)
            {
                Console.WriteLine(oneGroup.Name);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsCsomCreateTermSet(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");

            TermSet myTermSet = myTermGroup.CreateTermSet(
                                                "CsCsomTermSet", Guid.NewGuid(), 1033);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsCsomFindTermSets(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");

            spCtx.Load(myTermGroup, gs => gs.TermSets);
            spCtx.ExecuteQuery();

            foreach (TermSet oneTermSet in myTermGroup.TermSets)
            {
                Console.WriteLine(oneTermSet.Name);
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsCsomCreateTerm(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
            TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");

            Term myTerm = myTermSet.CreateTerm("CsCsomTerm", 1033, Guid.NewGuid());
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsCsomFindTerms(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
            TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");

            spCtx.Load(myTermSet, ts => ts.Terms);
            spCtx.ExecuteQuery();

            foreach (Term oneTerm in myTermSet.Terms)
            {
                Console.WriteLine(oneTerm.Name);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsCsomFindOneTerm(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
            TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");
            Term myTerm = myTermSet.Terms.GetByName("CsCsomTerm");

            spCtx.Load(myTerm);
            spCtx.ExecuteQuery();

            Console.WriteLine(myTerm.Name);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsCsomUpdateOneTerm(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
            TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");
            Term myTerm = myTermSet.Terms.GetByName("CsCsomTerm");

            myTerm.Name = "CsCsomTerm_Updated";
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsCsomDeleteOneTerm(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
            TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");
            Term myTerm = myTermSet.Terms.GetByName("CsCsomTerm");

            myTerm.DeleteObject();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsCsomFindTermSetAndTermById(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
            TermSet myTermSet = myTermStore.GetTermSet(
                                    new Guid("fdf6890f-5e8b-4d69-8a94-af92fdcebf30"));
            Term myTerm = myTermStore.GetTerm(
                                    new Guid("2f5352b2-a929-472e-9e8a-5d2b4c119bd3"));

            spCtx.Load(myTermSet);
            spCtx.Load(myTerm);
            spCtx.ExecuteQuery();

            Console.WriteLine(myTermSet.Name + " - " + myTerm.Name);
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsPnpcoreCreateTermGroup(ClientContext spCtx)
        {
            string termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ==";

            TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
            TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);

            TermGroup myTermGroup = myTermStore.CreateTermGroup("CsPnpcoreTermGroup");
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsPnpcoreCreateTermGroupEnsure(ClientContext spCtx)
        {
            TermGroup myTermGroup = spCtx.Site.EnsureTermGroup("CsPnpcoreTermGroupEns");
            Console.WriteLine(myTermGroup.Id);
        }
        //gavdcodeend 13

        //gavdcodebegin 16
        static void SpCsPnpcoreFindTermGroup(ClientContext spCtx)
        {
            TermGroup myTermGroup = spCtx.Site.GetTermGroupByName("CsPnpcoreTermGroupEns");
            Console.WriteLine(myTermGroup.Id);
        }
        //gavdcodeend 16

        //gavdcodebegin 14
        static void SpCsPnpcoreCreateTermSetEnsure(ClientContext spCtx)
        {
            TermGroup myTermGroup = spCtx.Site.EnsureTermGroup("CsPnpcoreTermGroupEns");
            TermSet myTermSet = myTermGroup.EnsureTermSet("CsPnpcoreTermSetEns");
            Console.WriteLine(myTermSet.Id);
        }
        //gavdcodeend 14

        //gavdcodebegin 17
        static void SpCsPnpcoreFindTermSet(ClientContext spCtx)
        {
            TermSetCollection myTermSet = spCtx.Site.GetTermSetsByName(
                                                                "CsPnpcoreTermSetEns");
            Console.WriteLine(myTermSet[0].Id);
        }
        //gavdcodeend 17

        //gavdcodebegin 15
        static void SpCsPnpcoreCreateTerm(ClientContext spCtx)
        {
            TermGroup myTermGroup = spCtx.Site.EnsureTermGroup("CsPnpcoreTermGroupEns");
            TermSet myTermSet = myTermGroup.EnsureTermSet("CsPnpcoreTermSetEns");
            Term myTerm = spCtx.Site.AddTermToTermset(myTermSet.Id, "CsPnpcoreTerm");
            Console.WriteLine(myTerm.Id);
        }
        //gavdcodeend 15

        //gavdcodebegin 18
        static void SpCsPnpcoreFindTerm(ClientContext spCtx)
        {
            TermSetCollection myTermSet = spCtx.Site.GetTermSetsByName(
                                                                "CsPnpcoreTermSetEns");
            Term myTerm = spCtx.Site.GetTermByName(myTermSet[0].Id, "CsPnpcoreTerm");
            Console.WriteLine(myTerm.Id);
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        static void SpCsPnpcoreExportTermStore(ClientContext spCtx)
        {
            List<string> myTermStoreExport = spCtx.Site.ExportAllTerms(true);
            foreach (string oneTerm in myTermStoreExport)
            {
                Console.WriteLine(oneTerm);
            }
        }
        //gavdcodeend 19

        //gavdcodebegin 20
        static void SpCsPnpcoreImportTermStore(ClientContext spCtx)
        {
            string[] myTerms = { "TermGroup01|TermSet01|Term01",
                                 "TermGroup01|TermSet01|Term02" };

            spCtx.Site.ImportTerms(myTerms, 1033);
        }
        //gavdcodeend 20

        //gavdcodebegin 21
        static void SpCsCsomGetResultsSearch(ClientContext spCtx)
        {
            KeywordQuery keywordQuery = new KeywordQuery(spCtx);
            keywordQuery.QueryText = "Team";
            SearchExecutor searchExecutor = new SearchExecutor(spCtx);
            ClientResult<ResultTableCollection> results =
                                        searchExecutor.ExecuteQuery(keywordQuery);
            spCtx.ExecuteQuery();

            foreach (var resultRow in results.Value[0].ResultRows)
            {
                Console.WriteLine(resultRow["Title"] + " - " +
                                        resultRow["Path"] + " - " + resultRow["Write"]);
            }
        }
        //gavdcodeend 21

        //gavdcodebegin 22
        static void SpCsRestResultsSearchGET(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri +
                                "/_api/search/query?querytext='team'";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 22

        //gavdcodebegin 23
        static void SpCsRestResultsSearchPOST(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = new
                {
                    __metadata = new
                    {
                        type =
                                "Microsoft.Office.Server.Search.REST.SearchRequest"
                    },
                    Querytext = "team",
                    RowLimit = 20,
                    ClientType = "ContentSearchRegular"
                };
                string endpointUrl = webUri + "/_api/search/query";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 23

        //gavdcodebegin 24
        static void SpCsCsomGetAllPropertiesUserProfile(ClientContext spCtx)
        {
            string myUser = "i:0#.f|membership|" +
                                        ConfigurationManager.AppSettings["spUserName"];
            PeopleManager myPeopleManager = new PeopleManager(spCtx);
            PersonProperties myUserProperties = myPeopleManager.GetPropertiesFor(myUser);
            spCtx.Load(myUserProperties, prop => prop.AccountName,
                                                    prop => prop.UserProfileProperties);
            spCtx.ExecuteQuery();

            foreach (var oneProperty in myUserProperties.UserProfileProperties)
            {
                Console.WriteLine(oneProperty.Key.ToString() + " - " +
                                                        oneProperty.Value.ToString());
            }
        }
        //gavdcodeend 24

        //gavdcodebegin 25
        static void SpCsCsomGetAllMyPropertiesUserProfile(ClientContext spCtx)
        {
            PeopleManager myPeopleManager = new PeopleManager(spCtx);
            PersonProperties myUserProperties = myPeopleManager.GetMyProperties();
            spCtx.Load(myUserProperties, prop => prop.AccountName,
                                                    prop => prop.UserProfileProperties);
            spCtx.ExecuteQuery();

            foreach (var oneProperty in myUserProperties.UserProfileProperties)
            {
                Console.WriteLine(oneProperty.Key.ToString() + " - " +
                                                        oneProperty.Value.ToString());
            }
        }
        //gavdcodeend 25

        //gavdcodebegin 26
        static void SpCsCsomGetPropertiesUserProfile(ClientContext spCtx)
        {
            string myUser = "i:0#.f|membership|" +
                                        ConfigurationManager.AppSettings["spUserName"];
            PeopleManager myPeopleManager = new PeopleManager(spCtx);
            string[] myProfPropertyNames = new string[]
                                                   { "Manager", "Department", "Title" };
            UserProfilePropertiesForUser myProfProperties =
                new UserProfilePropertiesForUser(spCtx, myUser, myProfPropertyNames);
            IEnumerable<string> myProfPropertyValues =
                myPeopleManager.GetUserProfilePropertiesFor(myProfProperties);

            spCtx.Load(myProfProperties);
            spCtx.ExecuteQuery();

            foreach (string oneValue in myProfPropertyValues)
            {
                Console.WriteLine(oneValue);
            }
        }
        //gavdcodeend 26

        //gavdcodebegin 27
        static void SpCsCsomUpdateOnePropertyUserProfile(ClientContext spCtx)
        {
            PeopleManager myPeopleManager = new PeopleManager(spCtx);
            PersonProperties myUserProperties = myPeopleManager.GetMyProperties();
            spCtx.Load(myUserProperties, prop => prop.AccountName);
            spCtx.ExecuteQuery();

            string newValue = "I am the administrator";
            myPeopleManager.SetSingleValueProfileProperty(
                    myUserProperties.AccountName, "AboutMe", newValue);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 27

        //gavdcodebegin 28
        static void SpCsCsomUpdateOneMultPropertyUserProfile(ClientContext spCtx)
        {
            PeopleManager myPeopleManager = new PeopleManager(spCtx);
            PersonProperties myUserProperties = myPeopleManager.GetMyProperties();
            spCtx.Load(myUserProperties, prop => prop.AccountName);
            spCtx.ExecuteQuery();

            List<string> mySkills = new List<string>();
            mySkills.Add("SharePoint");
            mySkills.Add("Windows");
            myPeopleManager.SetMultiValuedProfileProperty(
                                    myUserProperties.AccountName, "SPS-Skills", mySkills);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 28

        //gavdcodebegin 29
        static void SpCsRestGetAllPropertiesUserProfile(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                string myUser = "i%3A0%23.f%7Cmembership%7C" +
                     ConfigurationManager.AppSettings["spUserName"].Replace("@", "%40");
                object myPayload = null;
                string endpointUrl = webUri +
                      "/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v='" +
                      myUser + "'";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 29

        //gavdcodebegin 30
        static void SpCsRestGetAllMyPropertiesUserProfile(Uri webUri, string userName,
                                                                    string password)
        {
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri +
                      "/_api/sp.userprofiles.peoplemanager/getmyproperties";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 30

        //gavdcodebegin 31
        static void SpCsRestGetPropertiesUserProfile(Uri webUri, string userName,
                                                                    string password)
        {
            string myUser = "i%3A0%23.f%7Cmembership%7C" +
                 ConfigurationManager.AppSettings["spUserName"].Replace("@", "%40");
            using (SPHttpClient client = new SPHttpClient(webUri, userName, password))
            {
                object myPayload = null;
                string endpointUrl = webUri +
                      "/_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor" +
                      "(accountame=@v, propertyname='AboutMe')?@v='" + myUser + "'";
                var data = client.ExecuteJson(endpointUrl, HttpMethod.Get, myPayload);
                Console.WriteLine(data);
            }
        }
        //gavdcodeend 31

        //gavdcodebegin 32
        static void SpCsCsomGenerateWebSiteScript(ClientContext spCtx)
        {
            Tenant myTenant = new Tenant(spCtx);

            TenantSiteScriptSerializationInfo myInfo = new TenantSiteScriptSerializationInfo()
            {
                IncludeBranding = true,
                IncludeTheme = true,
                IncludeRegionalSettings = true,
                IncludeLinksToExportedItems = true,
                IncludeSiteExternalSharingCapability = true,
                IncludedLists = new[] { "Shared Documents", "Lists/TestList" }
            };

            var response = myTenant.GetSiteScriptFromSite(
                "https://[domain].sharepoint.com/sites/Test_Guitaca/", myInfo);

            spCtx.ExecuteQuery();

            Console.WriteLine(response.Value.JSON);
        }
        //gavdcodeend 32

        //gavdcodebegin 33
        static void SpCsCsomAddSiteScript(ClientContext spCtx)
        {
            string myScript = System.IO.File.ReadAllText
                                            (@"C:\Temporary\TestListSiteScript.json");

            Tenant myTenant = new Tenant(spCtx);

            TenantSiteScriptCreationInfo myInfo = new TenantSiteScriptCreationInfo()
            {
                Title = "CustomListFromSiteScript",
                Content = myScript,
                Description = "Creates a Custom List using CSOM"
            };

            var response = myTenant.CreateSiteScript(myInfo);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 33

        //gavdcodebegin 34
        static void SpCsCsomGetAllSiteScripts(ClientContext spCtx)
        {
            Tenant myTenant = new Tenant(spCtx);

            // ATTENTION: It doesn't work
            var response = myTenant.GetSiteScripts();

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 34

        //gavdcodebegin 35
        static void SpCsCsomUpdateSiteScript(ClientContext spCtx)
        {
            Tenant myTenant = new Tenant(spCtx);

            // ATTENTION: There is no TenantSiteScriptUpdateInfo object
            //TenantSiteScriptUpdateInfo myInfo = new TenantSiteScriptUpdateInfo()
            //{
            //    Title = "CustomListFromSiteScript",
            //    Description = "Creates a Custom List using CSOM updated"
            //};

            //var response = myTenant.UpdateSiteScript(myInfo);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 35

        //gavdcodebegin 36
        static void SpCsCsomDeleteSiteScript(ClientContext spCtx)
        {
            Guid myId = new Guid("da06b992-aeaf-439d-a73a-08905ae3e884");

            Tenant myTenant = new Tenant(spCtx);

            myTenant.DeleteSiteScript(myId);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 36

        //gavdcodebegin 37
        static void SpCsCsomAddSiteDesign(ClientContext spCtx)
        {
            Guid myId = new Guid("79a5174f-0712-49c7-b6af-5a45918c55ee");

            Tenant myTenant = new Tenant(spCtx);

            TenantSiteDesignCreationInfo myInfo = new TenantSiteDesignCreationInfo()
            {
                Title = "Custom List From Site Design CSOM",
                WebTemplate = "64",
                SiteScriptIds = new Guid[] { myId },
                Description = "Creates a Custom List in a site using CSOM Site Design"
            };

            var response = myTenant.CreateSiteDesign(myInfo);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 37

        //gavdcodebegin 38
        static void SpCsCsomApplySiteDesign(ClientContext spCtx)
        {
            string mySiteUrl = "https://[domain].sharepoint.com/sites/aaa05";
            Guid myId = new Guid("abed53c3-4515-4308-8821-ffc3ec3dbcdb");

            Tenant myTenant = new Tenant(spCtx);

            var response = myTenant.ApplySiteDesign(mySiteUrl, myId);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 38

        //gavdcodebegin 39
        static void SpCsCsomGetAllSiteDesigns(ClientContext spCtx)
        {
            Tenant myTenant = new Tenant(spCtx);

            // ATTENTION: It doesn't work
            var response = myTenant.GetSiteDesigns();

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 39

        //gavdcodebegin 40
        static void SpCsCsomUpdateSiteDesign(ClientContext spCtx)
        {
            Tenant myTenant = new Tenant(spCtx);

            // ATTENTION: There is no TenantSiteDesignUpdateInfo object
            //TenantSiteDesignUpdateInfo myInfo = new TenantSiteDesignUpdateInfo()
            //{
            //    Title = "CustomListFromSiteScript",
            //    Description = "Creates a Custom List using CSOM updated"
            //};

            //var response = myTenant.UpdateSiteDesign(myInfo);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 40

        //gavdcodebegin 41
        static void SpCsCsomGetTasksSiteDesign(ClientContext spCtx)
        {
            string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";

            Tenant myTenant = new Tenant(spCtx);

            var response = myTenant.GetSiteDesignTasks(mySiteUrl);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 41

        //gavdcodebegin 42
        static void SpCsCsomGetRunsSiteDesign(ClientContext spCtx)
        {
            string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";
            Guid myId = new Guid("79a5174f-0712-49c7-b6af-5a45918c55ee");

            Tenant myTenant = new Tenant(spCtx);

            var response = myTenant.GetSiteDesignRun(mySiteUrl, myId);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 42

        //gavdcodebegin 43
        static void SpCsCsomGetRunStatusSiteDesign(ClientContext spCtx)
        {
            string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";
            Guid myId = new Guid("79a5174f-0712-49c7-b6af-5a45918c55ee");

            Tenant myTenant = new Tenant(spCtx);

            var myRuns = myTenant.GetSiteDesignRun(mySiteUrl, myId);
            if (myRuns.AreItemsAvailable == true)
            {
                foreach (TenantSiteDesignRun oneRun in myRuns)
                {
                    var response = myTenant.GetSiteDesignRunStatus(oneRun.SiteId,
                                                                   oneRun.WebId,
                                                                   oneRun.Id);
                    Console.WriteLine(response.ToString());
                }
            }

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 43

        //gavdcodebegin 44
        static void SpCsCsomGrantRightsSiteDesign(ClientContext spCtx)
        {
            Guid myId = new Guid("da06b992-aeaf-439d-a73a-08905ae3e884");
            string[] myPrincipals = new string[] { "[user]@[domain].onmicrosoft.com" };
            TenantSiteDesignPrincipalRights myRights = TenantSiteDesignPrincipalRights.View;

            Tenant myTenant = new Tenant(spCtx);

            myTenant.GrantSiteDesignRights(myId, myPrincipals, myRights);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 44

        //gavdcodebegin 45
        static void SpCsCsomDeleteSiteDesign(ClientContext spCtx)
        {
            Guid myId = new Guid("abed53c3-4515-4308-8821-ffc3ec3dbcdb");

            Tenant myTenant = new Tenant(spCtx);

            myTenant.DeleteSiteDesign(myId);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 45

        //gavdcodebegin 46
        static void SpCsPnpcoreGenerateSiteTemplateXml(ClientContext spCtx)
        {
            OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.
                ProvisioningTemplateCreationInformation myProvisioner = 
                            new OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.
                                     ProvisioningTemplateCreationInformation(spCtx.Web);

            myProvisioner.FileConnector = 
                    new OfficeDevPnP.Core.Framework.Provisioning.Connectors.
                        FileSystemConnector(@"C:\Temporary\InMemTemplate.xml", "");

            // Write progress to output
            myProvisioner.ProgressDelegate = delegate (String prMessage, 
                                                       Int32 prProgress, 
                                                       Int32 prTotal)
            {
                Console.WriteLine("{0:00}/{1:00} - {2}", prProgress, prTotal, prMessage);
            };

            // Template in memory
            OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate 
                            myTemplate = spCtx.Web.GetProvisioningTemplate(myProvisioner);

            // Template in file
            OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.XMLTemplateProvider 
                myXmlProvider = new OfficeDevPnP.Core.Framework.Provisioning.Providers.
                    Xml.XMLFileSystemTemplateProvider(@"C:\Temporary", "");
            myXmlProvider.SaveAs(myTemplate, "TestProvisioningSite.xml");
        }
        //gavdcodeend 46

        //gavdcodebegin 47
        static void SpCsPnpcoreGenerateSiteListTemplate(ClientContext spCtx)
        {
            OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.
                ProvisioningTemplateCreationInformation myProvisioner = 
                            new OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.
                                     ProvisioningTemplateCreationInformation(spCtx.Web);

            myProvisioner.FileConnector = 
                    new OfficeDevPnP.Core.Framework.Provisioning.Connectors.
                        FileSystemConnector(@"C:\Temporary\InMemTemplate.xml", "");
            List<string> myLists = new List<string>();
            myLists.Add("MyCustomList");
            myLists.Add("Shared Documents");
            myProvisioner.ListsToExtract = myLists;

            // Write progress to output
            myProvisioner.ProgressDelegate = delegate (String prMessage, 
                                                       Int32 prProgress, 
                                                       Int32 prTotal)
            {
                Console.WriteLine("{0:00}/{1:00} - {2}", prProgress, prTotal, prMessage);
            };

            // Template in memory
            OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate 
                            myTemplate = spCtx.Web.GetProvisioningTemplate(myProvisioner);

            // Template in file
            OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.XMLTemplateProvider 
                myXmlProvider = new OfficeDevPnP.Core.Framework.Provisioning.Providers.
                    Xml.XMLFileSystemTemplateProvider(@"C:\Temporary", "");
            myXmlProvider.SaveAs(myTemplate, "TestProvisioningLists.xml");
        }
        //gavdcodeend 47

        //gavdcodebegin 48
        static void SpCsPnpcoreApplySiteTemplate(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            spCtx.Load(myWeb, w => w.Title);
            spCtx.ExecuteQueryRetry();

            OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.XMLTemplateProvider 
                myXmlProvider = new OfficeDevPnP.Core.Framework.Provisioning.Providers.
                    Xml.XMLFileSystemTemplateProvider(@"C:\Temporary", "");

            OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningTemplate 
                myTemplate = myXmlProvider.GetTemplate("TestProvisioningSite.xml");

            myWeb.ApplyProvisioningTemplate(myTemplate);
        }
        //gavdcodeend 48

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

        //-------------------------------------------------------------------------------
        static ClientContext LoginCsomAdminSite()
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

        //-------------------------------------------------------------------------------
        static ClientContext LoginPnPCore()
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
    }

    //-----------------------------------------------------------------------------------
    class SPHttpClientHandler : HttpClientHandler
    {
        public SPHttpClientHandler(Uri webUri, string userName, string password)
        {
            CookieContainer = GetAuthCookies(webUri, userName, password);
            FormatType = FormatType.JsonVerbose;
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,
                                                CancellationToken cancellationToken)
        {
            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            if (FormatType == FormatType.JsonVerbose)
            {
                request.Headers.Add("Accept", "application/json;odata=verbose");
            }
            return base.SendAsync(request, cancellationToken);
        }

        private static CookieContainer GetAuthCookies(Uri webUri,
                                                string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(userName, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(webUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(webUri, authCookie);
            return cookieContainer;
        }

        public FormatType FormatType { get; set; }
    }

    public enum FormatType
    {
        JsonVerbose,
        Xml
    }

    class SPHttpClient : HttpClient
    {
        public SPHttpClient(Uri webUri, string userName, string password) : base(
                                new SPHttpClientHandler(webUri, userName, password))
        {
            BaseAddress = webUri;
        }

        public object ExecuteJson(string requestUri, HttpMethod method,
                                    IDictionary<string, string> headers, object payload,
                                    bool GetBinaryResponse = false)
        {
            HttpResponseMessage response;
            switch (method.Method)
            {
                case "POST":
                    DefaultRequestHeaders.Add("X-RequestDigest", RequestFormDigest());
                    if (headers != null)
                    {
                        foreach (var header in headers)
                        {
                            DefaultRequestHeaders.Add(header.Key, header.Value);
                        }
                    }
                    if ((payload != null) && (payload.GetType().Name == "FileStream"))
                    {
                        StreamContent requestContent = new StreamContent((Stream)payload);
                        requestContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");
                        response = PostAsync(requestUri, requestContent).Result;
                    }
                    else
                    {
                        StringContent requestContent = null; // = new StringContent(string.Empty);
                        if (payload.GetType().FullName == "System.String")
                        {
                            requestContent = new StringContent((string)payload);
                        }
                        else
                        {
                            requestContent = new StringContent(
                                                   JsonConvert.SerializeObject(payload));
                        }
                        requestContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");
                        response = PostAsync(requestUri, requestContent).Result;
                    }
                    break;
                case "GET":
                    response = GetAsync(requestUri).Result;
                    break;
                default:
                    throw new NotSupportedException(string.Format(
                                        "Method {0} is not supported", method.Method));
            }

            //response.EnsureSuccessStatusCode();

            if (GetBinaryResponse == true)
            {
                var responseContentStream = response.Content.ReadAsStreamAsync().Result;
                return responseContentStream;
            }
            else
            {
                var responseContent = response.Content.ReadAsStringAsync().Result;
                return String.IsNullOrEmpty(responseContent) ? new JObject() :
                                                        JObject.Parse(responseContent);
            }
        }

        public object ExecuteJson<T>(string requestUri, HttpMethod method, T payload,
                                        bool GetBinaryResponse = false)
        {
            return ExecuteJson(requestUri, method, null, payload, GetBinaryResponse);
        }

        public object ExecuteJson(string requestUri, bool GetBinaryResponse = false)
        {
            return ExecuteJson(requestUri, HttpMethod.Get, null, default(string),
                                                                    GetBinaryResponse);
        }

        public string RequestFormDigest()
        {
            var endpointUrl = string.Format("{0}/_api/contextinfo", BaseAddress);
            var result = this.PostAsync(endpointUrl, new StringContent(
                                                            string.Empty)).Result;
            result.EnsureSuccessStatusCode();
            var content = result.Content.ReadAsStringAsync().Result;
            var contentJson = JObject.Parse(content);
            return contentJson["d"]["GetContextWebInformation"][
                                                        "FormDigestValue"].ToString();
        }
    }
}