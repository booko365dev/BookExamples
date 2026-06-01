using Microsoft.Identity.Client;
using System.Configuration;
using System.Text;
using System.Text.Json;
using System.Web;
using System.Xml;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 10.0 Console Application ****---------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static AuthenticationResult GetTokenWithDeviceCode(string clientId, string tenantId, string[] scopes)
{
    var app = PublicClientApplicationBuilder.Create(clientId)
        .WithTenantId(tenantId)
        .Build();

    return app.AcquireTokenWithDeviceCode(scopes, deviceCodeResult =>
    {
        Console.WriteLine(deviceCodeResult.Message);
        return Task.CompletedTask;
    }).ExecuteAsync().Result;
}

// The next two routines are for demonstration purposes only, and not recommended for production
//      use. Consider using the above routine with device code flow instead, or other more
//      secure authentication flows.
// Use instead the GetTokenWithDeviceCode routine to get an access token.
static Tuple<string, string> GetTokenWithAccPw()
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string myEndpoint = "https://login.microsoftonline.com/" +
                        ConfigurationManager.AppSettings["TenantName"] + "/oauth2/token";

    string reqBody = $"resource={ConfigurationManager.AppSettings["SiteBaseUrl"]}&";
    reqBody += $"grant_type=password&";
    reqBody += $"client_id=" +
                $"{ConfigurationManager.AppSettings["ClientIdWithAccPw"]}&";
    reqBody += $"username=" +
                $"{HttpUtility.UrlEncode(ConfigurationManager.AppSettings["UserName"])}&";
    reqBody += $"password=" +
                $"{HttpUtility.UrlEncode(ConfigurationManager.AppSettings["UserPw"])}";

    using (StringContent myStrContent = new(reqBody, Encoding.UTF8,
                                                    "application/x-www-form-urlencoded"))
    {
        HttpClient myHttpClient = new();
        string tokenStr = myHttpClient.PostAsync(myEndpoint,
                            myStrContent).ContinueWith((myResponse) =>
                            {
                                return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
                            }).Result;

        var tokenObj = JsonSerializer.Deserialize<JsonElement>(tokenStr);
        JsonElement myError;
        bool hasError = tokenObj.TryGetProperty("error", out myError);

        if (hasError == true)
        {
            string strError = "TokenErrorException - " +
                        tokenObj.GetProperty("error").GetString() + " - " +
                        tokenObj.GetProperty("error_description").GetString();

            tplReturn = new Tuple<string, string>(
                        tokenObj.GetProperty("error_codes")[0].GetRawText(), strError);
        }
        else
        {
            string myToken = tokenObj.GetProperty("access_token").GetString();

            tplReturn = new Tuple<string, string>("OK", myToken);
        }
    }

    return tplReturn;
}

static string GetRequestDigest(Tuple<string, string> AuthToken)
{
    string strReturn = string.Empty;
    Tuple<string, string> myTokenWithAccPw;

    if (AuthToken == null)
    {
        myTokenWithAccPw = GetTokenWithAccPw();
    }
    else
    {
        myTokenWithAccPw = AuthToken;
    }

    string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                "/_api/contextinfo";

    string myBody = "{}";

    using (StringContent myStrContent = new(myBody, Encoding.UTF8,
                                             "application/json"))
    {
        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                 "Authorization", "Bearer " + myTokenWithAccPw.Item2);

        string digestXml = myHttpClient.PostAsync(myEndpoint,
                            myStrContent).ContinueWith((myResponse) =>
                            {
                                return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
                            }).Result;

        XmlDocument myDocXml = new();
        myDocXml.LoadXml(digestXml);

        XmlNodeList allNodes = myDocXml.SelectNodes("/");
        strReturn = allNodes[0]["d:GetContextWebInformation"]["d:FormDigestValue"].InnerText;
    }

    return strReturn;
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpRest_FindAppCatalog()
{
    AuthenticationResult myAuth = GetTokenWithDeviceCode(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["TenantName"],
                            new string[] { ConfigurationManager.AppSettings["SiteBaseUrl"] + "/.default" });

    if (myAuth != null)
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/_api/SP_TenantSettings_Current";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myAuth.AccessToken);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
    else
    {
        Console.WriteLine("Authentication failed.");
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpRest_FindTenantProps()
{
    AuthenticationResult myAuth = GetTokenWithDeviceCode(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["TenantName"],
                            new string[] { ConfigurationManager.AppSettings["SiteBaseUrl"] + "/.default" });

    if (myAuth != null)
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/_api/web/AllProperties";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myAuth.AccessToken);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
    else
    {
        Console.WriteLine("Authentication failed.");
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpRest_FindTenantOneProp()
{
    AuthenticationResult myAuth = GetTokenWithDeviceCode(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["TenantName"],
                            new string[] { ConfigurationManager.AppSettings["SiteBaseUrl"] + "/.default" });

    if (myAuth != null)
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                          "/sites/appcatalog/_api/web/GetStorageEntity('[PropertyName]')";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myAuth.AccessToken);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
    else
    {
        Console.WriteLine("Authentication failed.");
    }
}
//gavdcodeend 003


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 003 ***

//CsSpRest_FindAppCatalog();
//CsSpRest_FindTenantProps();
//CsSpRest_FindTenantOneProp();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
