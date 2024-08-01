using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpRest_CreateOneList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

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
static void CsSpRest_ReadAllLists()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/lists?$select=Title,Id";

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
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpRest_ReadOneList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/lists/getbytitle('NewListRestCs')";

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
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpRest_UpdateOneList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                               "/_api/lists/getbytitle('NewListRestCs')";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.List" },
            Description = "New List Description"
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
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpRest_DeleteOneList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                             "/_api/lists/getbytitle('NewListRestCs')";

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
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpRest_AddOneFieldToList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('NewListRestCs')/fields";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.Field" },
            Title = "MyMultilineField",
            FieldTypeKind = 3
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
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpRest_ReadAllFieldsFromList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/lists/getbytitle('NewListRestCs')/fields";

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
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpRest_ReadOneFieldFromList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('NewListRestCs')/fields/getbytitle('MyMultilineField')";

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
static void CsSpRest_UpdateOneFieldInList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('NewListRestCs')/fields/getbytitle('MyMultilineField')";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.Field" },
            Description = "New Field Description"
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
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpRest_DeleteOneFieldFromList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
         "/_api/lists/getbytitle('NewListRestCs')/fields/getbytitle('MyMultilineField')";

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
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpRest_BreakSecurityInheritanceList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                 "/_api/lists/getbytitle('NewListRestCs')/" +
                 "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)";

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

            Console.WriteLine("Done");
        }
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpRest_ResetSecurityInheritanceList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                       "/_api/lists/getbytitle('NewListRestCs')/resetroleinheritance";

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
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpRest_AddUserToSecurityRoleInList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
             "/_api/web/siteusers?$select=Id&$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                 "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                          "Accept", "application/json; odata=verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        userId = int.Parse(resultObj["d"]["results"][0]["Id"].ToString());
        Console.WriteLine(userId);
    }

    // Find the RoleDefinitions
    int roleId = 0;
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
         "/_api/web/roledefinitions?$select=Id&$filter=startswith(Name,'Full Control')";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                 "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                        "Accept", "application/json; odata=verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        roleId = int.Parse(resultObj["d"]["results"][0]["Id"].ToString());
        Console.WriteLine(roleId);
    }

    // Add the User in the RoleDefinition to the List
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('NewListRestCs')/roleassignments/addroleassignment" +
          "(principalid=" + userId + ",roledefid=" + roleId + ")";

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

            Console.WriteLine("Done");
        }
    }
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpRest_UpdateUserSecurityRoleInList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
             "/_api/web/siteusers?$select=Id&$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                  "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                         "Accept", "application/json; odata=verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        userId = int.Parse(resultObj["d"]["results"][0]["Id"].ToString());
        Console.WriteLine(userId);
    }

    // Find the RoleDefinitions
    int roleId = 0;
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/web/roledefinitions/getbyname('Edit')/Id";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                   "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                         "Accept", "application/json; odata=verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        roleId = int.Parse(resultObj["d"]["Id"].ToString());
        Console.WriteLine(roleId);
    }

    // Add the User in the RoleDefinition to the List
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('NewListRestCs')/roleassignments/addroleassignment" +
          "(principalid=" + userId + ",roledefid=" + roleId + ")";

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

            Console.WriteLine("Done");
        }
    }
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpRest_DeleteUserFromSecurityRoleInList()
{
    Tuple<string, string> myTokenWithAccPw = CsSpRest_GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
            "/_api/web/siteusers?$select=Id&$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new();
        myHttpClient.DefaultRequestHeaders.Add(
                                  "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                         "Accept", "application/json; odata=verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        userId = int.Parse(resultObj["d"]["results"][0]["Id"].ToString());
        Console.WriteLine(userId);
    }

    // Remove the User from the List
    if (myTokenWithAccPw.Item1.Equals("ok", StringComparison.CurrentCultureIgnoreCase))
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('NewListRestCs')/roleassignments/" +
          "getbyprincipalid(principalid=" + userId + ")";

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
//gavdcodeend 015


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 015 ***

//CsSpRest_CreateOneList();
//CsSpRest_ReadAllLists();
//CsSpRest_ReadOneList();
//CsSpRest_UpdateOneList();
//CsSpRest_DeleteOneList();
//CsSpRest_AddOneFieldToList();
//CsSpRest_ReadAllFieldsFromList();
//CsSpRest_ReadOneFieldFromList();
//CsSpRest_UpdateOneFieldInList();
//CsSpRest_DeleteOneFieldFromList();
//CsSpRest_BreakSecurityInheritanceList();
//CsSpRest_ResetSecurityInheritanceList();
//CsSpRest_AddUserToSecurityRoleInList();
//CsSpRest_UpdateUserSecurityRoleInList();
//CsSpRest_DeleteUserFromSecurityRoleInList();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

