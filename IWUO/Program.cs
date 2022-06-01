using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
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

//gavdcodebegin 01
static void SpCsCsomCreateOneSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    string myUser = ConfigurationManager.AppSettings["UserName"];
    SiteCreationProperties mySiteCreationProps = new SiteCreationProperties
    {
        Url = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom01",
        Title = "NewSiteCollectionModernCsCsom01",
        Owner = ConfigurationManager.AppSettings["UserName"],
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
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsCsomRestoreSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RestoreDeletedSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsCsomRemoveDeletedSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RemoveDeletedSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsCsomCreateGroupForSite(ClientContext spAdminCtx)
{
    string[] myOwners = new string[] { "user@domain.onmicrosoft.com" };
    GroupCreationParams myGroupParams = new GroupCreationParams(spAdminCtx);
    myGroupParams.Owners = myOwners;

    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.CreateGroupForSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
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
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
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
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewHubSiteCollCsCsom");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsCsomUnregisterAsHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.UnregisterHubSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewHubSiteCollCsCsom");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsCsomGetHubSiteCollectionProperties(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    HubSiteProperties myProps = myTenant.GetHubSitePropertiesByUrl(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
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
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
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
            ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewSiteForHub",
            ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");
    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsCsomremoveSiteFromHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.DisconnectSiteFromHubSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
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
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;
        spCtx.Load(myWeb);
        spCtx.ExecuteQuery();

        Console.WriteLine(myWeb.Title + " - " + myWeb.Url + " - " + myWeb.Id);
    }
}
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsCsomUpdateOneWebInSiteCollection()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;
        myWeb.Description = "NewWebSiteModernCsCsom Description Updated";
        myWeb.Update();
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsCsomDeleteOneWebInSiteCollection()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;
        myWeb.DeleteObject();
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 19

//gavdcodebegin 20
static void SpCsCsomBreakSecurityInheritanceWeb()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
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
}
//gavdcodeend 20

//gavdcodebegin 21
static void SpCsCsomResetSecurityInheritanceWeb()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
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
}
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsCsomAddUserToSecurityRoleInWeb()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        RoleDefinitionBindingCollection roleDefinition =
                new RoleDefinitionBindingCollection(spCtx);
        roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
        myWeb.RoleAssignments.Add(myUser, roleDefinition);

        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsCsomUpdateUserSecurityRoleInWeb()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        RoleDefinitionBindingCollection roleDefinition =
                new RoleDefinitionBindingCollection(spCtx);
        roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Contributor));

        RoleAssignment myRoleAssignment = myWeb.RoleAssignments.GetByPrincipal(myUser);
        myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

        myRoleAssignment.Update();
        spCtx.ExecuteQuery();
    }
}
//gavdcodeend 23

//gavdcodebegin 24
static void SpCsCsomDeleteUserFromSecurityRoleInWeb()
{
    string myWebFullUrl = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                    "/NewWebSiteModernCsCsom";

    SecureString usrPw = new SecureString();
    foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
        usrPw.AppendChar(oneChar);

    using (AuthenticationManager myAuthenticationManager =
                new AuthenticationManager())
    using (ClientContext spCtx = myAuthenticationManager.GetContext(
                new Uri(myWebFullUrl),
                ConfigurationManager.AppSettings["UserName"],
                usrPw,
                ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
    {
        Web myWeb = spCtx.Web;

        User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
        myWeb.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

        spCtx.ExecuteQuery();
        spCtx.Dispose();
    }
}
//gavdcodeend 24


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

////--> Working with Site Collections
//SecureString usrPw = new SecureString();
//foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
//    usrPw.AppendChar(oneChar);

//using (AuthenticationManager myAuthenticationManager =
//            new AuthenticationManager())
//using (ClientContext spAdminCtx = myAuthenticationManager.GetContext(
//            new Uri(ConfigurationManager.AppSettings["SiteAdminUrl"]),
//            ConfigurationManager.AppSettings["UserName"],
//            usrPw,
//            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
//{
//SpCsCsomCreateOneSiteCollection(spAdminCtx);
//SpCsCsomCreateGroupForSite(spAdminCtx);
//SpCsCsomFindWebTemplates(spAdminCtx);
//SpCsCsomReadAllSiteCollections(spAdminCtx);
//SpCsCsomRemoveSiteCollection(spAdminCtx);
//SpCsCsomRestoreSiteCollection(spAdminCtx);
//SpCsCsomRemoveDeletedSiteCollection(spAdminCtx);
//SpCsCsomSetAdministratorSiteCollection(spAdminCtx);
//SpCsCsomRegisterAsHubSiteCollection(spAdminCtx);
//SpCsCsomUnregisterAsHubSiteCollection(spAdminCtx);
//SpCsCsomGetHubSiteCollectionProperties(spAdminCtx);
//SpCsCsomUpdateHubSiteCollectionProperties(spAdminCtx);
//SpCsCsomAddSiteToHubSiteCollection(spAdminCtx);
//SpCsCsomremoveSiteFromHubSiteCollection(spAdminCtx);

//Console.WriteLine("Done");
//}

////--> Working with Web Sites
//SecureString usrPw = new SecureString();
//foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
//    usrPw.AppendChar(oneChar);

//using (AuthenticationManager myAuthenticationManager =
//            new AuthenticationManager())
//using (ClientContext spCtx = myAuthenticationManager.GetContext(
//            new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]),
//            ConfigurationManager.AppSettings["UserName"],
//            usrPw,
//            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
//{
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
//SpCsCsomDeleteUserFromSecurityRoleInWeb();

//    Console.WriteLine("Done");
//}


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

