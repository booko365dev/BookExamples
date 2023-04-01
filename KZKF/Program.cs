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

// Using a Class, see at the end of the file


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void SpCsCsom_CreateOneList(ClientContext spCtx)
{
    ListCreationInformation myListCreationInfo = new ListCreationInformation();
    myListCreationInfo.Title = "NewListCsCsom";
    myListCreationInfo.Description = "New List created using CSharp CSOM";
    myListCreationInfo.TemplateType = (int)ListTemplateType.GenericList;

    List newList = spCtx.Web.Lists.Add(myListCreationInfo);
    newList.OnQuickLaunch = true;
    newList.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 001

//gavdcodebegin 002
static void SpCsCsom_ReadAllList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    ListCollection allLists = myWeb.Lists;
    spCtx.Load(allLists, lsts => lsts.Include(lst => lst.Title,
                                              lst => lst.Id));
    spCtx.ExecuteQuery();

    foreach (List oneList in allLists)
    {
        Console.WriteLine(oneList.Title + " - " + oneList.Id);
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void SpCsCsom_ReadOneList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    spCtx.Load(myList);
    spCtx.ExecuteQuery();

    Console.WriteLine("List description - " + myList.Description);
}
//gavdcodeend 003

//gavdcodebegin 004
static void SpCsCsom_UpdateOneList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    myList.Description = "New List Description";
    myList.Update();
    spCtx.Load(myList);
    spCtx.ExecuteQuery();

    Console.WriteLine("List description - " + myList.Description);
}
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsCsom_DeleteOneList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    myList.DeleteObject();
    spCtx.ExecuteQuery();
}
//gavdcodeend 005

//gavdcodebegin 06
static void SpCsCsom_AddOneFieldToList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    string fieldXml = "<Field DisplayName='MyMultilineField' Type='Note' />";
    Field myField = myList.Fields.AddFieldAsXml(fieldXml,
                                               true,
                                               AddFieldOptions.DefaultValue);
    spCtx.ExecuteQuery();
}
//gavdcodeend 06

//gavdcodebegin 007
static void SpCsCsom_ReadAllFieldsFromList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    FieldCollection allFields = myList.Fields;
    spCtx.Load(allFields, flds => flds.Include(fld => fld.Title,
                                               fld => fld.TypeAsString));
    spCtx.ExecuteQuery();

    foreach (Field oneField in allFields)
    {
        Console.WriteLine(oneField.Title + " - " + oneField.TypeAsString);
    }
}
//gavdcodeend 007

//gavdcodebegin 008
static void SpCsCsom_ReadOneFieldFromList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    Field myField = myList.Fields.GetByTitle("MyMultilineField");
    spCtx.Load(myField);
    spCtx.ExecuteQuery();

    Console.WriteLine(myField.Id + " - " + myField.TypeAsString);
}
//gavdcodeend 008

//gavdcodebegin 009
static void SpCsCsom_UpdateOneFieldInList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    Field myField = myList.Fields.GetByTitle("MyMultilineField");

    FieldMultiLineText myFieldNote = spCtx.CastTo<FieldMultiLineText>(myField);
    myFieldNote.Description = "New Field Description";
    myFieldNote.Hidden = false;
    myFieldNote.NumberOfLines = 3;

    myField.Update();
    spCtx.Load(myField);
    spCtx.ExecuteQuery();

    Console.WriteLine(myField.Description);
}
//gavdcodeend 009

//gavdcodebegin 010
static void SpCsCsom_DeleteOneFieldFromList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    Field myField = myList.Fields.GetByTitle("MyMultilineField");
    myField.DeleteObject();
    spCtx.ExecuteQuery();
}
//gavdcodeend 010

//gavdcodebegin 011
static void SpCsCsom_BreakSecurityInheritanceList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    spCtx.Load(myList, hura => hura.HasUniqueRoleAssignments);
    spCtx.ExecuteQuery();

    if (myList.HasUniqueRoleAssignments == false)
    {
        myList.BreakRoleInheritance(false, true);
    }
    myList.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 011

//gavdcodebegin 012
static void SpCsCsom_AddUserToSecurityRoleInList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition =
                                        new RoleDefinitionBindingCollection(spCtx);
    roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
    myList.RoleAssignments.Add(myUser, roleDefinition);

    spCtx.ExecuteQuery();
}
//gavdcodeend 012

//gavdcodebegin 013
static void SpCsCsom_UpdateUserSecurityRoleInList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    RoleDefinitionBindingCollection roleDefinition =
                                        new RoleDefinitionBindingCollection(spCtx);
    roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Contributor));

    RoleAssignment myRoleAssignment = myList.RoleAssignments.GetByPrincipal(myUser);
    myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

    myRoleAssignment.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 013

//gavdcodebegin 014
static void SpCsCsom_DeleteUserFromSecurityRoleInList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

    User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["UserName"]);
    myList.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

    spCtx.ExecuteQuery();
}
//gavdcodeend 014

//gavdcodebegin 015
static void SpCsCsom_ResetSecurityInheritanceList(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
    spCtx.Load(myList, hura => hura.HasUniqueRoleAssignments);
    spCtx.ExecuteQuery();

    if (myList.HasUniqueRoleAssignments == true)
    {
        myList.ResetRoleInheritance();
    }
    myList.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 015

//gavdcodebegin 016
static void SpCsCsom_FieldCreateText(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

    Guid myGuid = Guid.NewGuid();
    string schemaField = "<Field ID='" + myGuid + "' Type='Text' " +
        "Name='myTextCol' StaticName='myTextCol' DisplayName='My Text Col' />";
    Field myField = myList.Fields.AddFieldAsXml(schemaField, true,
                AddFieldOptions.AddFieldInternalNameHint |
                AddFieldOptions.AddToDefaultContentType);

    spCtx.ExecuteQuery();
}
//gavdcodeend 016

//gavdcodebegin 17
static void SpCsCsom_ReadAllSiteColumns(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    FieldCollection allSiteColls = myWeb.Fields;

    spCtx.Load(allSiteColls, flds => flds.Include(fld => fld.Title,
                                               fld => fld.Group));
    spCtx.ExecuteQuery();

    foreach (Field oneColl in allSiteColls)
    {
        Console.WriteLine(oneColl.Title + " - " + oneColl.Group);
    }
}
//gavdcodeend 17

//gavdcodebegin 18
static void SpCsCsom_AddOneSiteColumn(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;

    string fieldXml = "<Field DisplayName='MySiteColMultilineField' " +
                                            "Type='Note' Group='MyGroup' />";
    Field myField = myWeb.Fields.AddFieldAsXml(fieldXml,
                                               true,
                                               AddFieldOptions.DefaultValue);
    spCtx.ExecuteQuery();
}
//gavdcodeend 18

//gavdcodebegin 19
static void SpCsCsom_ColumnIndex(ClientContext spCtx)
{
    Web myWeb = spCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

    string myColumn = "My Text Col";
    Field myField = myList.Fields.GetByTitle(myColumn);
    myField.Indexed = true;
    myField.Update();

    spCtx.ExecuteQuery();
}
//gavdcodeend 19


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//--> Working with Lists and Libraries
SecureString usrPw = new SecureString();
foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
    usrPw.AppendChar(oneChar);

using (AuthenticationManager myAuthenticationManager =
            new AuthenticationManager())
using (ClientContext spCtx = myAuthenticationManager.GetContext(
            new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]),
            ConfigurationManager.AppSettings["UserName"],
            usrPw,
            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
{
    //SpCsCsom_CreateOneList(spCtx);
    //SpCsCsom_ReadAllList(spCtx);
    //SpCsCsom_ReadOneList(spCtx);
    //SpCsCsom_UpdateOneList(spCtx);
    //SpCsCsom_DeleteOneList(spCtx);
    //SpCsCsom_AddOneFieldToList(spCtx);
    //SpCsCsom_ReadAllFieldsFromList(spCtx);
    //SpCsCsom_ReadOneFieldFromList(spCtx);
    //SpCsCsom_UpdateOneFieldInList(spCtx);
    //SpCsCsom_DeleteOneFieldFromList(spCtx);
    //SpCsCsom_BreakSecurityInheritanceList(spCtx);
    //SpCsCsom_AddUserToSecurityRoleInList(spCtx);
    //SpCsCsom_UpdateUserSecurityRoleInList(spCtx);
    //SpCsCsom_DeleteUserFromSecurityRoleInList(spCtx);
    //SpCsCsom_ResetSecurityInheritanceList(spCtx);
    //SpCsCsom_FieldCreateText(spCtx);
    //SpCsCsom_ReadAllSiteColumns(spCtx);
    //SpCsCsom_AddOneSiteColumn(spCtx);
    //SpCsCsom_ColumnIndex(spCtx);

    Console.WriteLine("Done");
}


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

