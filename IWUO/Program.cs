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

//gavdcodebegin 001
static void SpCsCsom_CreateOneSiteCollection(ClientContext spAdminCtx)
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
//gavdcodeend 001

//gavdcodebegin 002
static void SpCsCsom_FindWebTemplates(ClientContext spAdminCtx)
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
//gavdcodeend 002

//gavdcodebegin 003
static void SpCsCsom_ReadAllSiteCollections(ClientContext spAdminCtx)
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
//gavdcodeend 003

//gavdcodebegin 004
static void SpCsCsom_RemoveSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RemoveSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsCsom_RestoreSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RestoreDeletedSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 005

//gavdcodebegin 006
static void SpCsCsom_RemoveDeletedSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RemoveDeletedSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom02");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 006

//gavdcodebegin 007
static void SpCsCsom_CreateGroupForSite(ClientContext spAdminCtx)
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
//gavdcodeend 007

//gavdcodebegin 008
static void SpCsCsom_SetAdministratorSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.SetSiteAdmin(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/sites/NewSiteCollectionModernCsCsom01",
                                        "user@domain.onmicrosoft.com",
                                        true);

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 008

//gavdcodebegin 009
static void SpCsCsom_RegisterAsHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.RegisterHubSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewHubSiteCollCsCsom");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 009

//gavdcodebegin 010
static void SpCsCsom_UnregisterAsHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.UnregisterHubSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewHubSiteCollCsCsom");

    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 010

//gavdcodebegin 011
static void SpCsCsom_GetHubSiteCollectionProperties(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    HubSiteProperties myProps = myTenant.GetHubSitePropertiesByUrl(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewHubSiteCollCsCsom");

    spAdminCtx.Load(myProps);
    spAdminCtx.ExecuteQuery();

    Console.WriteLine(myProps.Title);
}
//gavdcodeend 011

//gavdcodebegin 012
static void SpCsCsom_UpdateHubSiteCollectionProperties(ClientContext spAdminCtx)
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
//gavdcodeend 012

//gavdcodebegin 013
static void SpCsCsom_AddSiteToHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.ConnectSiteToHubSite(
            ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewSiteForHub",
            ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                             "/sites/NewHubSiteCollCsCsom");
    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 013

//gavdcodebegin 014
static void SpCsCsom_RemoveSiteFromHubSiteCollection(ClientContext spAdminCtx)
{
    Tenant myTenant = new Tenant(spAdminCtx);
    myTenant.DisconnectSiteFromHubSite(
        ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                     "/sites/NewSiteForHub");
    spAdminCtx.ExecuteQuery();
}
//gavdcodeend 014

//gavdcodebegin 015
static void SpCsCsom_CreateOneWebInSiteCollection(ClientContext spCtx)
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
//gavdcodeend 015

//gavdcodebegin 016
static void SpCsCsom_GetWebsInSiteCollection(ClientContext spCtx)
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
//gavdcodeend 016

//gavdcodebegin 017
static void SpCsCsom_GetOneWebInSiteCollection()
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
//gavdcodeend 017

//gavdcodebegin 018
static void SpCsCsom_UpdateOneWebInSiteCollection()
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
//gavdcodeend 018

//gavdcodebegin 019
static void SpCsCsom_DeleteOneWebInSiteCollection()
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
//gavdcodeend 019

//gavdcodebegin 020
static void SpCsCsom_BreakSecurityInheritanceWeb()
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
//gavdcodeend 020

//gavdcodebegin 021
static void SpCsCsom_ResetSecurityInheritanceWeb()
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
//gavdcodeend 021

//gavdcodebegin 022
static void SpCsCsom_AddUserToSecurityRoleInWeb()
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
//gavdcodeend 022

//gavdcodebegin 023
static void SpCsCsom_UpdateUserSecurityRoleInWeb()
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
//gavdcodeend 023

//gavdcodebegin 024
static void SpCsCsom_DeleteUserFromSecurityRoleInWeb()
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
//gavdcodeend 024


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
//SpCsCsom_CreateOneSiteCollection(spAdminCtx);
//SpCsCsom_CreateGroupForSite(spAdminCtx);
//SpCsCsom_FindWebTemplates(spAdminCtx);
//SpCsCsom_ReadAllSiteCollections(spAdminCtx);
//SpCsCsom_RemoveSiteCollection(spAdminCtx);
//SpCsCsom_RestoreSiteCollection(spAdminCtx);
//SpCsCsom_RemoveDeletedSiteCollection(spAdminCtx);
//SpCsCsom_SetAdministratorSiteCollection(spAdminCtx);
//SpCsCsom_RegisterAsHubSiteCollection(spAdminCtx);
//SpCsCsom_UnregisterAsHubSiteCollection(spAdminCtx);
//SpCsCsom_GetHubSiteCollectionProperties(spAdminCtx);
//SpCsCsom_UpdateHubSiteCollectionProperties(spAdminCtx);
//SpCsCsom_AddSiteToHubSiteCollection(spAdminCtx);
//SpCsCsom_removeSiteFromHubSiteCollection(spAdminCtx);

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
//SpCsCsom_CreateOneWebInSiteCollection(spCtx);
//SpCsCsom_GetWebsInSiteCollection(spCtx);
//SpCsCsom_GetOneWebInSiteCollection();
//SpCsCsom_UpdateOneWebInSiteCollection();
//SpCsCsom_DeleteOneWebInSiteCollection();
//SpCsCsom_BreakSecurityInheritanceWeb();
//SpCsCsom_ResetSecurityInheritanceWeb();
//SpCsCsom_AddUserToSecurityRoleInWeb();
//SpCsCsom_UpdateUserSecurityRoleInWeb();
//SpCsCsom_DeleteUserFromSecurityRoleInWeb();
//SpCsCsom_DeleteUserFromSecurityRoleInWeb();

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

