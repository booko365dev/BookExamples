using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using System.Configuration;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------




//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpCsom_GetPropertiesTenant(ClientContext spCtx)
{
    Tenant myTenant = new(spCtx);

    foreach (PropertyInfo oneProperty in myTenant.GetType().GetProperties())
    {
        Console.WriteLine(oneProperty.Name);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpCsom_GetValuePropertyTenant(ClientContext spCtx)
{
    Tenant myTenant = new(spCtx);

    spCtx.Load(myTenant);
    spCtx.ExecuteQuery();

    bool myAccessDevices = myTenant.BlockAccessOnUnmanagedDevices;
    Console.WriteLine(myAccessDevices);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpCsom_UpdateValuePropertyTenant(ClientContext spCtx)
{
    Tenant myTenant = new(spCtx)
    {
        BlockAccessOnUnmanagedDevices = true
    };
    myTenant.Update();
    spCtx.ExecuteQuery();
}
//gavdcodeend 003


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 003 ***

SecureString usrPw = new();
foreach (char oneChar in ConfigurationManager.AppSettings["UserPw"])
    usrPw.AppendChar(oneChar);

using (AuthenticationManager authenticationManager = new())
using (ClientContext spCtx = authenticationManager.GetContext(
            new Uri(ConfigurationManager.AppSettings["SiteAdminUrl"]),
            ConfigurationManager.AppSettings["UserName"],
            usrPw,
            ConfigurationManager.AppSettings["ClientIdWithAccPw"]))
{
    //CsSpCsom_GetPropertiesTenant(spCtx);
    //CsSpCsom_GetValuePropertyTenant(spCtx);
    //CsSpCsom_UpdateValuePropertyTenant(spCtx);

    Console.WriteLine("Done");
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

public class AuthenticationManager : IDisposable
{
    private static readonly HttpClient httpClient = new();
    private const string tokenEndpoint =
                            "https://login.microsoftonline.com/common/oauth2/token";

    private static readonly SemaphoreSlim semaphoreSlimTokens = new(1);
    private AutoResetEvent tokenResetEvent = null;
    private readonly ConcurrentDictionary<string, string> tokenCache = new();
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

                tokenResetEvent = new(false);
                TokenWaitInfo wi = new();
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
                            catch (Exception)
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

    private static async Task<string> AcquireTokenAsync(Uri resourceUri,
                                        string username, string password, string clientId)
    {
        string resource = $"{resourceUri.Scheme}://{resourceUri.DnsSafeHost}";

        var body = $"resource={resource}&";
        body += $"client_id={clientId}&";
        body += $"grant_type=password&";
        body += $"username={HttpUtility.UrlEncode(username)}&";
        body += $"password={HttpUtility.UrlEncode(password)}";
        using var stringContent = new StringContent(body,
                            Encoding.UTF8, "application/x-www-form-urlencoded");
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
#pragma warning restore CS8321 // Local function is declared but never used

