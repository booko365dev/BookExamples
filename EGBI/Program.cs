using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using System.Collections.Concurrent;
using System.Configuration;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------




//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------




//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

SecureString usrPw = new SecureString();
foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
    usrPw.AppendChar(oneChar);

using (AuthenticationManager authenticationManager =
            new AuthenticationManager())
using (ClientContext spCtx = authenticationManager.GetContext(
            new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]),
            ConfigurationManager.AppSettings["UserName"],
            usrPw,
            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
{
    // CSOM Term Store
    //SpCsCsom_FindTermStore(spCtx);
    //SpCsCsom_CreateTermGroup(spCtx);
    //SpCsCsom_FindTermGroups(spCtx);
    //SpCsCsom_CreateTermSet(spCtx);
    //SpCsCsom_FindTermSets(spCtx);
    //SpCsCsom_CreateTerm(spCtx);
    //SpCsCsom_FindTerms(spCtx);
    //SpCsCsom_FindOneTerm(spCtx);
    //SpCsCsom_UpdateOneTerm(spCtx);
    //SpCsCsom_DeleteOneTerm(spCtx);
    //SpCsCsom_FindTermSetAndTermById(spCtx);

    // Search
    //SpCsCsom_GetResultsSearch(spCtx);

    // UserProfile
    //SpCsCsom_GetAllPropertiesUserProfile(spCtx);
    //SpCsCsom_GetAllMyPropertiesUserProfile(spCtx);
    //SpCsCsom_GetPropertiesUserProfile(spCtx);
    //SpCsCsom_UpdateOnePropertyUserProfile(spCtx);
    //SpCsCsom_UpdateOneMultPropertyUserProfile(spCtx);

    // Site Scripts
    //SpCsCsom_GenerateWebSiteScript(spCtx);
    //SpCsCsom_AddSiteScript(spCtx);
    //SpCsCsom_GetAllSiteScripts(spCtx);
    //SpCsCsom_UpdateSiteScript(spCtx);

    // Site Templates
    //SpCsCsom_AddSiteTemplate(spCtx);
    //SpCsCsom_ApplySiteTemplate(spCtx);
    //SpCsCsom_GetAllSiteTemplates(spCtx);
    //SpCsCsom_UpdateSiteTemplate(spCtx);
    //SpCsCsom_GetTasksSiteTemplate(spCtx);
    //SpCsCsom_GetRunsSiteTemplate(spCtx);
    //SpCsCsom_GetRunStatusSiteTemplate(spCtx);
    //SpCsCsom_GrantRightsSiteTemplate(spCtx);
    //SpCsCsom_DeleteSiteTemplate(spCtx);

    Console.WriteLine("Done");
}

// --- CSOM Term Store
//gavdcodebegin 01
static void SpCsCsom_FindTermStore(ClientContext spCtx)
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
static void SpCsCsom_CreateTermGroup(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

    TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
    TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);

    TermGroup myTermGroup = myTermStore.CreateGroup(
                                            "CsCsomTermGroup", Guid.NewGuid());
    spCtx.ExecuteQuery();
}
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsCsom_FindTermGroups(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_CreateTermSet(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

    TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
    TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
    TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");

    TermSet myTermSet = myTermGroup.CreateTermSet(
                                        "CsCsomTermSet", Guid.NewGuid(), 1033);
    spCtx.ExecuteQuery();
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsCsom_FindTermSets(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_CreateTerm(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

    TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
    TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
    TermGroup myTermGroup = myTermStore.Groups.GetByName("CsCsomTermGroup");
    TermSet myTermSet = myTermGroup.TermSets.GetByName("CsCsomTermSet");

    Term myTerm = myTermSet.CreateTerm("CsCsomTerm", 1033, Guid.NewGuid());
    spCtx.ExecuteQuery();
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsCsom_FindTerms(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_FindOneTerm(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_UpdateOneTerm(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_DeleteOneTerm(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

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
static void SpCsCsom_FindTermSetAndTermById(ClientContext spCtx)
{
    string termStoreName = "Taxonomy_A16ApXAPRyrML/PibplHbA==";

    TaxonomySession myTaxSession = TaxonomySession.GetTaxonomySession(spCtx);
    TermStore myTermStore = myTaxSession.TermStores.GetByName(termStoreName);
    TermSet myTermSet = myTermStore.GetTermSet(
                            new Guid("7d40eadb-c320-4320-8eb0-da725c8a426f"));
    Term myTerm = myTermStore.GetTerm(
                            new Guid("8279e7c6-6508-48f4-b8fc-456710a4f6b8"));

    spCtx.Load(myTermSet);
    spCtx.Load(myTerm);
    spCtx.ExecuteQuery();

    Console.WriteLine(myTermSet.Name + " - " + myTerm.Name);
}
//gavdcodeend 11

// --- Search
//gavdcodebegin 12
static void SpCsCsom_GetResultsSearch(ClientContext spCtx)
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
//gavdcodeend 12

// --- User Profile
//gavdcodebegin 13
static void SpCsCsom_GetAllPropertiesUserProfile(ClientContext spCtx)
{
    string myUser = "i:0#.f|membership|" +
                                ConfigurationManager.AppSettings["UserName"];
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
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsCsom_GetAllMyPropertiesUserProfile(ClientContext spCtx)
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
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsCsom_GetPropertiesUserProfile(ClientContext spCtx)
{
    string myUser = "i:0#.f|membership|" +
                                ConfigurationManager.AppSettings["UserName"];
    PeopleManager myPeopleManager = new PeopleManager(spCtx);
    string[] myProfPropertyNames = new string[]
                                           { "Manager", "Department", "WorkEmail" };
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
//gavdcodeend 15

//gavdcodebegin 16
static void SpCsCsom_UpdateOnePropertyUserProfile(ClientContext spCtx)
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
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsCsom_UpdateOneMultPropertyUserProfile(ClientContext spCtx)
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
//gavdcodeend 17

// -- Site Scripts
//gavdcodebegin 18
static void SpCsCsom_GenerateWebSiteScript(ClientContext spCtx)
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

    ClientResult<TenantSiteScriptSerializationResult> response = 
        myTenant.GetSiteScriptFromSite(
                    ConfigurationManager.AppSettings["SiteCollUrl"], 
                    myInfo);

    spCtx.ExecuteQuery();

    Console.WriteLine(response.Value.JSON);
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsCsom_AddSiteScript(ClientContext spCtx)
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

    TenantSiteScript myScriptResult = myTenant.CreateSiteScript(myInfo);

    spCtx.ExecuteQuery();
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsCsom_GetAllSiteScripts(ClientContext spCtx)
{
    Tenant myTenant = new Tenant(spCtx);

    ClientObjectList<TenantSiteScript> myScript = myTenant.GetSiteScripts();

    spCtx.ExecuteQuery();
}
//gavdcodeend 20

//gavdcodebegin 21
static void SpCsCsom_UpdateSiteScript(ClientContext spCtx)
{
    Tenant myTenant = new Tenant(spCtx);

    TenantSiteScript myInfo = new TenantSiteScript(spCtx, null)
    {
        Title = "CustomListFromSiteScript",
        Description = "Creates a Custom List using CSOM updated"
    };

    TenantSiteScript myScript = myTenant.UpdateSiteScript(myInfo);

    spCtx.ExecuteQuery();
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsCsom_DeleteSiteScript(ClientContext spCtx)
{
    Guid myId = new Guid("da06b992-aeaf-439d-a73a-08905ae3e884");

    Tenant myTenant = new Tenant(spCtx);

    myTenant.DeleteSiteScript(myId);

    spCtx.ExecuteQuery();
}
//gavdcodeend 22

// -- Site Templates
//gavdcodebegin 23
static void SpCsCsom_AddSiteTemplate(ClientContext spCtx)
{
    Guid myId = new Guid("79a5174f-0712-49c7-b6af-5a45918c55ee");

    Tenant myTenant = new Tenant(spCtx);

    TenantSiteDesignCreationInfo myInfo = new TenantSiteDesignCreationInfo()
    {
        Title = "Custom List From Site Template CSOM",
        WebTemplate = "64",
        SiteScriptIds = new Guid[] { myId },
        Description = "Creates a Custom List in a site using CSOM Site Template"
    };

    var response = myTenant.CreateSiteDesign(myInfo);

    spCtx.ExecuteQuery();
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsCsom_ApplySiteTemplate(ClientContext spCtx)
{
    string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";
    Guid myId = new Guid("abed53c3-4515-4308-8821-ffc3ec3dbcdb");

    Tenant myTenant = new Tenant(spCtx);

    var response = myTenant.ApplySiteDesign(mySiteUrl, myId);

    spCtx.ExecuteQuery();
}
//gavdcodeend 24

//gavdcodebegin 25
static void SpCsCsom_GetAllSiteTemplates(ClientContext spCtx)
{
    Tenant myTenant = new Tenant(spCtx);

    var response = myTenant.GetSiteDesigns();

    spCtx.ExecuteQuery();
}
//gavdcodeend 25

//gavdcodebegin 26
static void SpCsCsom_UpdateSiteTemplate(ClientContext spCtx)
{
    Tenant myTenant = new Tenant(spCtx);

    TenantSiteDesign myInfo = new TenantSiteDesign(spCtx, null)
    {
        Title = "CustomListFromSiteScript",
        Description = "Creates a Custom List using CSOM updated"
    };

    var response = myTenant.UpdateSiteDesign(myInfo);

    spCtx.ExecuteQuery();
}
//gavdcodeend 26

//gavdcodebegin 27
static void SpCsCsom_GetTasksSiteTemplate(ClientContext spCtx)
{
    string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";

    Tenant myTenant = new Tenant(spCtx);

    var response = myTenant.GetSiteDesignTasks(mySiteUrl);

    spCtx.ExecuteQuery();
}
//gavdcodeend 27

//gavdcodebegin 28
static void SpCsCsom_GetRunsSiteTemplate(ClientContext spCtx)
{
    string mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca";
    Guid myId = new Guid("79a5174f-0712-49c7-b6af-5a45918c55ee");

    Tenant myTenant = new Tenant(spCtx);

    var response = myTenant.GetSiteDesignRun(mySiteUrl, myId);

    spCtx.ExecuteQuery();
}
//gavdcodeend 28

//gavdcodebegin 29
static void SpCsCsom_GetRunStatusSiteTemplate(ClientContext spCtx)
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
//gavdcodeend 29

//gavdcodebegin 30
static void SpCsCsom_GrantRightsSiteTemplate(ClientContext spCtx)
{
    Guid myId = new Guid("da06b992-aeaf-439d-a73a-08905ae3e884");
    string[] myPrincipals = new string[] { "[user]@[domain].onmicrosoft.com" };
    TenantSiteDesignPrincipalRights myRights = TenantSiteDesignPrincipalRights.View;

    Tenant myTenant = new Tenant(spCtx);

    myTenant.GrantSiteDesignRights(myId, myPrincipals, myRights);

    spCtx.ExecuteQuery();
}
//gavdcodeend 30

//gavdcodebegin 31
static void SpCsCsom_DeleteSiteTemplate(ClientContext spCtx)
{
    Guid myId = new Guid("abed53c3-4515-4308-8821-ffc3ec3dbcdb");

    Tenant myTenant = new Tenant(spCtx);

    myTenant.DeleteSiteDesign(myId);

    spCtx.ExecuteQuery();
}
//gavdcodeend 31


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

public class AuthenticationManager : IDisposable
{
    private static readonly HttpClient httpClient = new HttpClient();
    private const string tokenEndpoint =
                            "https://login.microsoftonline.com/common/oauth2/token";

    private static readonly SemaphoreSlim semaphoreSlimTokens = new SemaphoreSlim(1);
    private AutoResetEvent tokenResetEvent = null;
    private readonly ConcurrentDictionary<string, string> tokenCache =
                                            new ConcurrentDictionary<string, string>();
    private bool disposedValue;

    internal class TokenWaitInfo
    {
        public RegisteredWaitHandle Handle = null;
    }

    public ClientContext GetContext(Uri web, string userPrincipalName,
                                            SecureString userPassword, string clientId)
    {
        var context = new ClientContext(web);

        context.ExecutingWebRequest += (sender, e) =>
        {
            string accessToken = EnsureAccessTokenAsync(
               new Uri($"{web.Scheme}://{web.DnsSafeHost}"),
               userPrincipalName,
               new System.Net.NetworkCredential(string.Empty, userPassword).Password,
               clientId).GetAwaiter().GetResult();

            if (accessToken.Contains("TokenErrorException") == true)
            {
                throw new Exception(accessToken); // An error has been raised by AAD
            }

            e.WebRequestExecutor.RequestHeaders["Authorization"] =
                "Bearer " + accessToken;
        };

        return context;
    }

    public async Task<string> EnsureAccessTokenAsync(Uri resourceUri,
                        string userPrincipalName, string userPassword, string clientId)
    {
        string accessTokenFromCache = TokenFromCache(resourceUri, tokenCache);
        if (accessTokenFromCache == null)
        {
            await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
            try
            {
                string accessToken = await AcquireTokenAsync(resourceUri,
                    userPrincipalName, userPassword, clientId).ConfigureAwait(false);

                if (accessToken.Contains("TokenErrorException") == true)
                { return accessToken; } // An error has been raised by Azure AD

                AddTokenToCache(resourceUri, tokenCache, accessToken);

                tokenResetEvent = new AutoResetEvent(false);
                TokenWaitInfo wi = new TokenWaitInfo();
                wi.Handle = ThreadPool.RegisterWaitForSingleObject(
                    tokenResetEvent,
                    async (state, timedOut) =>
                    {
                        if (!timedOut)
                        {
                            TokenWaitInfo wi1 = (TokenWaitInfo)state;
                            if (wi1.Handle != null)
                            {
                                wi1.Handle.Unregister(null);
                            }
                        }
                        else
                        {
                            try
                            {
                                await semaphoreSlimTokens.WaitAsync().
                                                            ConfigureAwait(false);
                                RemoveTokenFromCache(resourceUri, tokenCache);
                            }
                            catch (Exception ex)
                            {
                                RemoveTokenFromCache(resourceUri, tokenCache);
                            }
                            finally
                            {
                                semaphoreSlimTokens.Release();
                            }
                        }
                    },
                    wi,
                    (uint)CalculateThreadSleep(accessToken).TotalMilliseconds,
                    true
                );

                return accessToken;
            }
            finally
            {
                semaphoreSlimTokens.Release();
            }
        }
        else
        {
            return accessTokenFromCache;
        }
    }

    private async Task<string> AcquireTokenAsync(Uri resourceUri,
                                        string username, string password, string clientId)
    {
        string resource = $"{resourceUri.Scheme}://{resourceUri.DnsSafeHost}";

        var body = $"resource={resource}&";
        body += $"client_id={clientId}&";
        body += $"grant_type=password&";
        body += $"username={HttpUtility.UrlEncode(username)}&";
        body += $"password={HttpUtility.UrlEncode(password)}";
        using (var stringContent = new StringContent(body,
                            Encoding.UTF8, "application/x-www-form-urlencoded"))
        {
            var result = await httpClient.PostAsync(tokenEndpoint,
                            stringContent).ContinueWith((response) =>
                            {
                                return response.Result.Content.ReadAsStringAsync().Result;
                            }).ConfigureAwait(false);

            var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
            try
            { // Check for an error returned by Azure AD
                var tokenError = tokenResult.GetProperty("error").GetString();

                string strError = "TokenErrorException - " +
                            tokenResult.GetProperty("error").GetString() + " - " +
                            tokenResult.GetProperty("error_description").GetString();

                return strError;
            }
            catch
            { } // Nothing to catch, the response is giving correctly the token 

            var token = tokenResult.GetProperty("access_token").GetString();
            return token;
        }
    }

    private static string TokenFromCache(Uri web, ConcurrentDictionary<string,
                                                                    string> tokenCache)
    {
        if (tokenCache.TryGetValue(web.DnsSafeHost, out string accessToken))
        {
            return accessToken;
        }

        return null;
    }

    private static void AddTokenToCache(Uri web, ConcurrentDictionary<string,
                                            string> tokenCache, string newAccessToken)
    {
        if (tokenCache.TryGetValue(web.DnsSafeHost, out string currentAccessToken))
        {
            tokenCache.TryUpdate(web.DnsSafeHost, newAccessToken, currentAccessToken);
        }
        else
        {
            tokenCache.TryAdd(web.DnsSafeHost, newAccessToken);
        }
    }

    private static void RemoveTokenFromCache(Uri web, ConcurrentDictionary<string,
                                                                    string> tokenCache)
    {
        tokenCache.TryRemove(web.DnsSafeHost, out string currentAccessToken);
    }

    private static TimeSpan CalculateThreadSleep(string accessToken)
    {
        var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
        var lease = GetAccessTokenLease(token.ValidTo);
        lease = TimeSpan.FromSeconds(lease.TotalSeconds -
            TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds -
            TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
        return lease;
    }

    private static TimeSpan GetAccessTokenLease(DateTime expiresOn)
    {
        DateTime now = DateTime.UtcNow;
        DateTime expires = expiresOn.Kind == DateTimeKind.Utc ? expiresOn :
            TimeZoneInfo.ConvertTimeToUtc(expiresOn);
        TimeSpan lease = expires - now;
        return lease;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                if (tokenResetEvent != null)
                {
                    tokenResetEvent.Set();
                    tokenResetEvent.Dispose();
                }
            }

            disposedValue = true;
        }
    }

    // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method  
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}

#nullable enable
