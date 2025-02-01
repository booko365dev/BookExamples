using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Configuration;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Web;
using System.Xml;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static Tuple<string, string> CsSpSharePointRest_GetTokenWithAccPw_Rest()
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
                        return myResponse.Result.Content.ReadAsStringAsync().Result;
                    }).Result;

        var tokenObj = System.Text.Json.JsonSerializer.Deserialize<JsonElement>(tokenStr);
        bool hasError = tokenObj.TryGetProperty("error", out JsonElement myError);

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
//gavdcodeend 001

//gavdcodebegin 007
static Tuple<string, string> CsSpSharePointRest_GetTokenWithAccPw_Msal()
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string tenantName = ConfigurationManager.AppSettings["TenantName"];
    string clientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string userName = ConfigurationManager.AppSettings["UserName"];
    string userPw = ConfigurationManager.AppSettings["UserPw"];
    string siteBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"]; ;
    string myAuthority = $"https://login.microsoftonline.com/{tenantName}";

    IPublicClientApplication myApp = PublicClientApplicationBuilder
        .Create(clientId)
        .WithAuthority(new Uri(myAuthority))
        .Build();

    string[] myScopes = [$"{siteBaseUrl}/.default"];

    try
    {
        AuthenticationResult myResult = myApp
            .AcquireTokenByUsernamePassword(myScopes, userName, userPw)
            .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 007

static Tuple<string, string> CsSpSharePointRest_GetTokenWithSecret_Rest()
{
    // ATENTION: This method is not working, because the SharePoint REST API does not
    //      accept authentication of App Registrations with secrets. The method returns
    //      a token, but the token is not accepted by the SharePoint REST API

    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string myEndpoint = "https://login.microsoftonline.com/" +
                        ConfigurationManager.AppSettings["TenantName"] + "/oauth2/token";

    string siteBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"];
    string clientId = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];

    string reqBody = $"resource={siteBaseUrl}&";
    reqBody += $"grant_type=client_credentials&";
    reqBody += $"client_id=" +
                $"{clientId}&";
    reqBody += $"client_secret=" +
                $"{HttpUtility.UrlEncode(clientSecret)}";

    using (StringContent myStrContent = new(reqBody, Encoding.UTF8,
                                                    "application/x-www-form-urlencoded"))
    {
        HttpClient myHttpClient = new();
        string tokenStr = myHttpClient.PostAsync(myEndpoint,
                    myStrContent).ContinueWith((myResponse) =>
                    {
                        return myResponse.Result.Content.ReadAsStringAsync().Result;
                    }).Result;

        var tokenObj = System.Text.Json.JsonSerializer.Deserialize<JsonElement>(tokenStr);
        bool hasError = tokenObj.TryGetProperty("error", out JsonElement myError);

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

static Tuple<string, string> CsSpSharePointRest_GetTokenWithSecret_Msal()
{
    // ATENTION: This method is not working, because the SharePoint REST API does not
    //      accept authentication of App Registrations with secrets. The method returns
    //      a token, but the token is not accepted by the SharePoint REST API

    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string tenantName = ConfigurationManager.AppSettings["TenantName"];
    string clientId = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
    string siteBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"]; ;
    string myAuthority = $"https://login.microsoftonline.com/{tenantName}";

    IConfidentialClientApplication myApp = ConfidentialClientApplicationBuilder
        .Create(clientId)
        .WithAuthority(new Uri(myAuthority))
        .WithClientSecret(clientSecret)
        .Build();

    string[] myScopes = [$"{siteBaseUrl}/.default"];

    try
    {
        AuthenticationResult myResult = myApp
            .AcquireTokenForClient(myScopes)
            .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}

//gavdcodebegin 008
static Tuple<string, string> CsSpSharePointRest_GetTokenWithCert_Msal()
{
    Tuple<string, string> tplReturn = new(string.Empty, string.Empty);

    string tenantName = ConfigurationManager.AppSettings["TenantName"];
    string clientId = ConfigurationManager.AppSettings["ClientIdWithCert"];
    string certificateFilePath = ConfigurationManager.AppSettings["CertificateFilePath"];
    string certificateFilePw = ConfigurationManager.AppSettings["CertificateFilePw"];
    string siteBaseUrl = ConfigurationManager.AppSettings["SiteBaseUrl"]; ;
    string myAuthority = $"https://login.microsoftonline.com/{tenantName}";

    X509Certificate2 certificate = new(certificateFilePath, certificateFilePw);

    IConfidentialClientApplication myApp = ConfidentialClientApplicationBuilder
        .Create(clientId)
        .WithAuthority(new Uri(myAuthority))
        .WithCertificate(certificate)
        .Build();

    string[] myScopes = [$"{siteBaseUrl}/.default"];

    try
    {
        AuthenticationResult myResult = myApp
            .AcquireTokenForClient(myScopes)
            .ExecuteAsync().Result;
        tplReturn = new Tuple<string, string>("OK", myResult.AccessToken);
    }
    catch (MsalServiceException ex)
    {
        string strError = "TokenErrorException - " + ex.ErrorCode + " - " + ex.Message;
        tplReturn = new Tuple<string, string>(ex.ErrorCode, strError);
    }

    return tplReturn;
}
//gavdcodeend 008

//gavdcodebegin 002
static string CsSpSharePointRest_GetRequestDigest(Tuple<string, string> AuthToken)
{
    string strReturn = string.Empty;
    Tuple<string, string> myToken;

    if (AuthToken == null)
    {
        myToken = CsSpSharePointRest_GetTokenWithAccPw_Msal();
    }
    else
    {
        myToken = AuthToken;
    }

    string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                "/_api/contextinfo";

    string myBody = "{}";

    using (StringContent myStrContent = new(myBody, Encoding.UTF8, "application/json"))
    {
        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                 "Authorization", "Bearer " + myToken.Item2);

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
//gavdcodeend 002


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 003
static void CsSpSharePointRest_TestSpRestGet()
{
    Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Msal();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Rest();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithCert_Msal();

    if (myToken.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/lists";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                "Authorization", "Bearer " + myToken.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
                        {
                            return myResponse.Result.Content.ReadAsStringAsync().Result;
                        }).Result;

        // Reading the query myResult, but only if the myResult is a JSON string
        dynamic resultObj = JsonConvert.DeserializeObject(resultStr);
        try
        {
            string strError = resultObj["odata.error"].code.Value;
            Console.WriteLine("Error found - " +
                                        resultObj["odata.error"].message.value.Value);
        }
        catch
        {
            try
            {
                string strOk = resultObj["odata.metadata"];

                foreach (var oneItem in resultObj["value"])
                {
                    Console.WriteLine(oneItem.Title.Value);
                }
            }
            catch
            {
                Console.WriteLine("Unknown error");
            }
        }
    }
    else
    {
        Console.WriteLine(myToken.Item2);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpSharePointRest_TestSpRestPost()
{
    Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Msal();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Rest();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithCert_Msal();

    if (myToken.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/lists";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.List" },
            Title = "NewListRestCs",
            BaseTemplate = 100,
            Description = "Test NewListRestCs",
            AllowContentTypes = true,
            ContentTypesEnabled = true
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                "Authorization", "Bearer " + myToken.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                "X-RequestDigest", CsSpSharePointRest_GetRequestDigest(myToken));

            string resultStr = myHttpClient.PostAsync(myEndpoint,
                                myStrContent).ContinueWith((myResponse) =>
                                {
                                    return myResponse.Result.Content
                                                            .ReadAsStringAsync().Result;
                                }).Result;

            var resultObj = System.Text.Json.JsonSerializer
                                        .Deserialize<JsonElement>(resultStr);
            bool hasError = resultObj.TryGetProperty("error", out JsonElement myError);

            if (hasError == true)
            {
                Console.WriteLine("QueryException - " + myError);
            }
            else
            {
                Console.WriteLine("Done - " + resultStr);
            }
        }
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpSharePointRest_TestSpRestUpdate()
{
    Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Msal();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Rest();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithCert_Msal();

    if (myToken.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/lists/getbytitle('NewListRestCs')";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.List" },
            Description = "Test NewListRestCs Updated"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                            "Authorization", "Bearer " + myToken.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                            "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                            "X-RequestDigest", CsSpSharePointRest_GetRequestDigest(null));
            myHttpClient.DefaultRequestHeaders.Add(
                            "IF-MATCH", "*");
            myHttpClient.DefaultRequestHeaders.Add(
                            "X-HTTP-Method", "MERGE");

            string resultStr = myHttpClient.PostAsync(myEndpoint,
                                myStrContent).ContinueWith((myResponse) =>
                                {
                                    return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
                                }).Result;

            if (resultStr != String.Empty)
            {
                var resultObj = System.Text.Json.JsonSerializer
                                                .Deserialize<JsonElement>(resultStr);
                Console.WriteLine("QueryException - " + resultObj.GetProperty("error"));
            }
            else
            {
                Console.WriteLine("Done");
            }
        }
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpSharePointRest_TestSpRestDelete()
{
    Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Msal();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithAccPw_Rest();
    //Tuple<string, string> myToken = CsSpSharePointRest_GetTokenWithCert_Msal();

    if (myToken.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/lists/getbytitle('NewListRestCs')";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                "Authorization", "Bearer " + myToken.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                "X-RequestDigest", CsSpSharePointRest_GetRequestDigest(myToken));
            myHttpClient.DefaultRequestHeaders.Add(
                "IF-MATCH", "*");
            myHttpClient.DefaultRequestHeaders.Add(
                "X-HTTP-Method", "DELETE");

            string resultStr = myHttpClient.PostAsync(myEndpoint,
                            myStrContent).ContinueWith((myResponse) =>
                            {
                                return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
                            }).Result;

            if (resultStr != String.Empty)
            {
                var resultObj = System.Text.Json.JsonSerializer
                                                .Deserialize<JsonElement>(resultStr);
                Console.WriteLine("QueryException - " + resultObj.GetProperty("error"));
            }
            else
            {
                Console.WriteLine("Done");
            }
        }
    }
}
//gavdcodeend 006


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 008 ***

//CsSpSharePointRest_TestSpRestGet();
//CsSpSharePointRest_TestSpRestPost();
//CsSpSharePointRest_TestSpRestUpdate();
//CsSpSharePointRest_TestSpRestDelete();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
