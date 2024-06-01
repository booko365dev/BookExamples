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

//gavdcodebegin 001
static Tuple<string, string> CsSpSharePointRest_GetTokenWithAccPw()
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
//gavdcodeend 001

//gavdcodebegin 002
static string CsSpSharePointRest_GetRequestDigest(Tuple<string, string> AuthToken)
{
    string strReturn = string.Empty;
    Tuple<string, string> myTokenWithAccPw;

    if (AuthToken == null)
    {
        myTokenWithAccPw = CsSpSharePointRest_GetTokenWithAccPw();
    }
    else
    {
        myTokenWithAccPw = AuthToken;
    }

    string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                "/_api/contextinfo";

    string myBody = "{}";

    using (StringContent myStrContent = new(myBody, Encoding.UTF8, "application/json"))
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
//gavdcodeend 002


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 003
static void CsSpSharePointRest_TestSpRestGet()
{
    Tuple<string, string> myTokenWithAccPw = CsSpSharePointRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/lists";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
                        {
                            return myResponse.Result.Content.ReadAsStringAsync().Result;
                        }).Result;

        // Reading the query result, but only if the result is a JSON string
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
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpSharePointRest_TestSpRestPost()
{
    Tuple<string, string> myTokenWithAccPw = CsSpSharePointRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
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
                "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                "X-RequestDigest", CsSpSharePointRest_GetRequestDigest(myTokenWithAccPw));

            string resultStr = myHttpClient.PostAsync(myEndpoint,
                                myStrContent).ContinueWith((myResponse) =>
                                {
                                    return myResponse.Result.Content
                                                            .ReadAsStringAsync().Result;
                                }).Result;

            var resultObj = System.Text.Json.JsonSerializer
                                        .Deserialize<JsonElement>(resultStr);
            JsonElement myError;
            bool hasError = resultObj.TryGetProperty("error", out myError);

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
    Tuple<string, string> myTokenWithAccPw = CsSpSharePointRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
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
                            "Authorization", "Bearer " + myTokenWithAccPw.Item2);
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
    Tuple<string, string> myTokenWithAccPw = CsSpSharePointRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
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
                "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                "X-RequestDigest", CsSpSharePointRest_GetRequestDigest(myTokenWithAccPw));
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

//# *** Latest Source Code Index: 006 ***

//CsSpSharePointRest_TestSpRestGet();
//CsSpSharePointRest_TestSpRestPost();
//CsSpSharePointRest_TestSpRestUpdate();
//CsSpSharePointRest_TestSpRestDelete();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
