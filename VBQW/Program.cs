using Newtonsoft.Json;
using System.Configuration;
using System.Net.Http.Headers;
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

static Tuple<string, string> CsSpRest_GetTokenWithAccPw()
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

static string GetRequestDigest(Tuple<string, string> AuthToken)
{
    string strReturn = string.Empty;
    Tuple<string, string> myTokenWithAccPw;

    if (AuthToken == null)
    {
        myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();
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
static void CsSpRest_CreateOneCommunicationSiteCollection()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/_api/sitepages/communicationsite/create";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.Publishing.CommunicationSiteCreationRequest" },
            Title = "NewSiteCollectionModernCsRest01",
            Description = "NewSiteCollectionModernCsRest Description",
            AllowFileSharingForGuestUsers = false,
            SiteDesignId = "6142d2a0-63a5-4ba0-aede-d9fefca2c767",
            Url = ConfigurationManager.AppSettings["SiteBaseUrl"] + 
                                                "/sites/NewSiteCollectionModernCsRest01",
            lcid = 1033
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                   "X-RequestDigest", GetRequestDigest(myTokenWithAccPw));

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
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpRest_CreateOneSiteCollection()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                        "/_api/SPSiteManager/create";

        object myPayloadObj = new
        {
            __metadata = new {type = "Microsoft.SharePoint.Portal.SPSiteCreationRequest"},
            Title = "NewSiteCollectionModernCsRest02",
            Description = "NewSiteCollectionModernCsRest Description",
            AllowFileSharingForGuestUsers = false,
            SiteDesignId = "00000000-0000-0000-0000-000000000000",
            WebTemplate = "SITEPAGEPUBLISHING#0",
            WebTemplateExtensionId = "00000000-0000-0000-0000-000000000000",
            Url = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                                "/sites/NewSiteCollectionModernCsRest02",
            lcid = 1033
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                   "X-RequestDigest", GetRequestDigest(myTokenWithAccPw));

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
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpRest_CreateOneWebInSiteCollection()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/webinfos/add";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.WebCreationInformation" },
            Title = "NewWebSiteModernCsRest",
            Description = "NewWebSiteModernCsRest Description",
            Url = "NewWebSiteModernCsRest",
            UseSamePermissionsAsParentSite = true,
            WebTemplate = "STS#3"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                   "X-RequestDigest", GetRequestDigest(myTokenWithAccPw));

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
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpRest_GetAllSiteCollections()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                                   "/_api/search/query?querytext='contentclass:sts_site'";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpRest_GetAllWebsInSiteCollection()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                                        "/_api/web/webs";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpRest_UpdateOneWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                      "/NewWebSiteModernCsRest/_api/web";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.Web" },
            Description = "NewWebSiteModernCsRest Description Updated"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                     "X-RequestDigest", GetRequestDigest(null));
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
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpRest_DeleteOneWebFromSiteCollection()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                       "/NewWebSiteModernCsRest/_api/web";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new();
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                   "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                   "X-RequestDigest", GetRequestDigest(myTokenWithAccPw));
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
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpRest_GetRoleDefinitionsWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                    "/NewWebSiteModernCsRest/_api/web/roledefinitions";


        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpRest_GetUserPermissionsWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                           "/NewWebSiteModernCsRest/_api/web/" +
                                            "doesuserhavepermissions(@v)?@v=" +
                                            "{'High':'2147483647', 'Low':'4294967295'}";


        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpRest_GetOtherUserPermissionsWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                           "/NewWebSiteModernCsRest/_api/web/" +
                                           "getusereffectivepermissions(@v)?@v=" +
                                           "'i%3A0%23.f%7Cmembership%7C" + "Admin" + "'";


        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        Console.WriteLine(resultStr);
    }
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpRest_BreakSecurityInheritanceWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                      "/NewWebSiteModernCsRest/_api/web" +
                                      "/breakroleinheritance(copyRoleAssignments=false," +
                                      "clearSubscopes=true)";

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
                                     "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                     "X-RequestDigest", GetRequestDigest(null));
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

            Console.WriteLine(resultStr);
        }
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpRest_ResetSecurityInheritanceWeb()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                  "/NewWebSiteModernCsRest/_api/web/resetroleinheritance";

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
                                     "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                     "X-RequestDigest", GetRequestDigest(null));
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

            Console.WriteLine(resultStr);
        }
    }
}
//gavdcodeend 012


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 012 ***

//CsSpRest_CreateOneCommunicationSiteCollection();
//CsSpRest_CreateOneSiteCollection();
//CsSpRest_CreateOneWebInSiteCollection();
//CsSpRest_GetAllSiteCollections();
//CsSpRest_GetAllWebsInSiteCollection();
//CsSpRest_UpdateOneWeb();
//CsSpRest_DeleteOneWebFromSiteCollection();
//CsSpRest_GetOtherUserPermissionsWeb();
//CsSpRest_GetUserPermissionsWeb();
//CsSpRest_GetRoleDefinitionsWeb();
//CsSpRest_BreakSecurityInheritanceWeb();
//CsSpRest_ResetSecurityInheritanceWeb();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
