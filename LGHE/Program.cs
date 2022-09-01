using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Web;
using System.Xml;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static Tuple<string, string> GetTokenWithAccPw()
{
    Tuple<string, string> tplReturn = new Tuple<string, string>(string.Empty, string.Empty);

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

    using (StringContent myStrContent = new StringContent(reqBody, Encoding.UTF8,
                                                    "application/x-www-form-urlencoded"))
    {
        HttpClient myHttpClient = new HttpClient();
        string tokenStr = myHttpClient.PostAsync(myEndpoint,
                            myStrContent).ContinueWith((myResponse) =>
                            {
                                return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
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

    using (StringContent myStrContent = new StringContent(myBody, Encoding.UTF8,
                                             "application/json"))
    {
        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                 "Authorization", "Bearer " + myTokenWithAccPw.Item2);

        string digestXml = myHttpClient.PostAsync(myEndpoint,
                            myStrContent).ContinueWith((myResponse) =>
                            {
                                return myResponse.Result.Content
                                                        .ReadAsStringAsync().Result;
                            }).Result;

        XmlDocument myDocXml = new XmlDocument();
        myDocXml.LoadXml(digestXml);

        XmlNodeList allNodes = myDocXml.SelectNodes("/");
        strReturn = allNodes[0]["d:GetContextWebInformation"]["d:FormDigestValue"].InnerText;
    }

    return strReturn;
}

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 01
static void SpCsRest_CreateOneListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/web/lists/getbytitle('TestList')/items";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "NewListItemCsRest"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 01

//gavdcodebegin 02
static void SpCsRest_UploadOneDocument()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        FileInfo myFileInfo = new FileInfo(@"C:\Temporary\TestText.txt");
        string webUrlRel = new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]).
                                                                            AbsolutePath;

        Stream myPayloadStream = System.IO.File.OpenRead(myFileInfo.FullName);
        StreamReader myPayloadReader = new StreamReader(myPayloadStream);
        string myPayload = myPayloadReader.ReadToEnd();

        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                "/_api/web/getfolderbyserverrelativeurl(" +
                                "'" + webUrlRel + "/TestLibrary')/files/add(url='" +
                                myFileInfo.Name + "',overwrite=true)";

        string myPayLoadJson = JsonConvert.SerializeObject(myPayload);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 02

//gavdcodebegin 03
static void SpCsRest_DownloadOneDocumentt()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string webUrlRel = new Uri(ConfigurationManager.AppSettings["SiteCollUrl"]).
                                                                            AbsolutePath;

        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                    "/_api/web/getfilebyserverrelativeurl(" +
                                    "'" + webUrlRel + "/TestLibrary/TestText.txt')" +
                                    "/$value";

        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        byte[] resultByte = Encoding.UTF8.GetBytes(resultStr);
        FileStream outputStream = new FileStream(@"C:\Temporary\TestText.txt",
                            FileMode.OpenOrCreate | FileMode.Append,
                            FileAccess.Write, FileShare.None);
        outputStream.Write(resultByte, 0, resultByte.Length);
        outputStream.Flush(true);
        outputStream.Close();
    }
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 03

//gavdcodebegin 04
static void SpCsRest_ReadAllListsItems()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                           "/_api/lists/getbytitle('TestList')/items?$select=Title,Id";

        HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 04

//gavdcodebegin 05
static void SpCsRest_ReadOneListsItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                "/_api/lists/getbytitle('TestList')" +
                                                "/items(13)?$select=Title,Id";

        HttpClient myHttpClient = new HttpClient();
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
        Console.WriteLine(resultObj.Title.Value);
    }
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 05

//gavdcodebegin 06
static void SpCsRest_ReadAllLibraryDocs()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                "/_api/lists/getbytitle('TestLibrary')" +
                                                "/items?$select=Title,Id";

        HttpClient myHttpClient = new HttpClient();
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
                    Console.WriteLine(oneItem.Id.Value);
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
//gavdcodeend 06

//gavdcodebegin 07
static void SpCsRest_ReadOneLibraryDoc()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                "/_api/lists/getbytitle('TestLibrary')" +
                                                "/items(10)?$select=Title,Id";

        HttpClient myHttpClient = new HttpClient();
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
        Console.WriteLine(resultObj.Id.Value);
    }
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 07

//gavdcodebegin 08
static void SpCsRest_UpdateOneListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestList')/items(13)";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "NewListItemCsRest_Updated"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 08

//gavdcodebegin 09
static void SpCsRest_UpdateOneLibraryDoc()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestLibrary')/items(10)";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Name = "NewDocCsRest_Updated"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 09

//gavdcodebegin 10
static void SpCsRest_DeleteOneListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestList')/items(15)";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 10

//gavdcodebegin 11
static void SpCsRest_DeleteOneLibraryDoc()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestLibrary')/items(20)";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 11

//gavdcodebegin 12
static void SpCsRest_BreakSecurityInheritanceListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                        "/_api/lists/getbytitle('TestList')/" +
                        "items(13)/breakroleinheritance(copyRoleAssignments=false," +
                        "clearSubscopes=true)";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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

            Console.WriteLine("Done");
        }
    }
}
//gavdcodeend 12

//gavdcodebegin 13
static void SpCsRest_ResetSecurityInheritanceListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestList')/" +
                                        "items(13)/resetroleinheritance";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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

            Console.WriteLine("Done");
        }
    }
}
//gavdcodeend 13

//gavdcodebegin 14
static void SpCsRest_AddUserToSecurityRoleInListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                       "Accept", "application/json; odata = verbose"); // Output as XML

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
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/web/roledefinitions?$select=Id&$filter=startswith(Name,'Full Control')";

        HttpClient myHttpClient = new HttpClient();
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

    // Add the User in the RoleDefinion to the ListItem
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('TestList')/items(13)" + 
          "/roleassignments/addroleassignment" +
          "(principalid=" + userId + ",roledefid=" + roleId + ")";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 14

//gavdcodebegin 15
static void SpCsRest_UpdateUserSecurityRoleInListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                       "Accept", "application/json; odata = verbose"); // Output as XML

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
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/roledefinitions/getbyname('Edit')/Id";

        HttpClient myHttpClient = new HttpClient();
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

    // Add the User in the RoleDefinion to the ListItem
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('TestList')/items(13)" + 
          "/roleassignments/addroleassignment" +
          "(principalid=" + userId + ",roledefid=" + roleId + ")";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 15

//gavdcodebegin 16
static void SpCsRest_DeleteUserFromSecurityRoleInListItem()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    // Find the User
    int userId = 0;
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/siteusers?$select=Id&" +
                                            "$filter=startswith(Title,'System Admin')";

        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                       "Accept", "application/json; odata = verbose"); // Output as XML

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        JObject resultObj = JObject.Parse(resultStr);
        userId = int.Parse(resultObj["d"]["results"][0]["Id"].ToString());
        Console.WriteLine(userId);
    }

    // Delete the User from the ListItem
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
          "/_api/lists/getbytitle('TestList')/items(13)" +
          "/roleassignments/getbyprincipalid(principalid=" + userId + ")";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Authorization", "Bearer " + myTokenWithAccPw.Item2);
            myHttpClient.DefaultRequestHeaders.Add(
                                     "Accept", "application/json;odata=verbose");
            myHttpClient.DefaultRequestHeaders.Add(
                                     "X-RequestDigest", GetRequestDigest(null));
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
//gavdcodeend 16

//gavdcodebegin 17
static void SpCsRest_CreateOneFolderInLibrary()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/[Site]/[Library]/FolderLibraryRest";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/web/Folders";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.Folder" },
            ServerRelativeUrl = myServerRelativeUrl
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 17

//gavdcodebegin 25
static void SpCsRest_CreateOneFolderInList()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    int itemId = 0;
    // Create the ListItem with Folder as ContentType
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/web/lists/getbytitle('TestList')/items";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "NewListItemCsRestForFolder",
            ContentTypeId = "0x0120"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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

            JObject resultObjA = JObject.Parse(resultStr);
            itemId = int.Parse(resultObjA["d"]["Id"].ToString());
        }
    }

    // Modify the properties of the Folder
    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/getbytitle('TestList')/items(" + 
                                        itemId + ")";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "CsRestFolder",
            FileLeafRef = "CsRestFolder"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 25

//gavdcodebegin 18
static void SpCsRest_ReadAllFoldersInLibrary()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/{Site]/[Library]";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/Web/GetFolderByServerRelativePath(" + 
                                   "decodedurl='" + myServerRelativeUrl + "')";

        HttpClient myHttpClient = new HttpClient();
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
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 18

//gavdcodebegin 26
static void SpCsRest_ReadAllFoldersInList()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                    "/_api/Web/Lists/getByTitle('TestList')/Items?" +
                                    "$filter=FSObjType eq '1'";
        // FSObjType == 0 --> File,  FSObjType == 1 --> Folder

        HttpClient myHttpClient = new HttpClient();
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
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 26

//gavdcodebegin 19
static void SpCsRest_RenameOneFolderInLibrary()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/[Site]/[Library]/RestFolder";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/GetFolderByServerRelativeUrl('" +
                                        myServerRelativeUrl + "')/ListItemAllFields";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "RestFolderRenamed",
            FileLeafRef = "RestFolderRenamed"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 19

//gavdcodebegin 27
static void SpCsRest_RenameOneFolderInList()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/[Site]/lists/[List]/RestFolder";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/GetFolderByServerRelativeUrl('" +
                                        myServerRelativeUrl + "')/ListItemAllFields";

        object myPayloadObj = new
        {
            __metadata = new { type = "SP.ListItem" },
            Title = "RestFolderRenamed",
            FileLeafRef = "RestFolderRenamed"
        };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 27

//gavdcodebegin 20
static void SpCsRest_DeleteOneFolderInLibrary()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/[Site]/[Library]/RestFolderRenamed";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/GetFolderByServerRelativeUrl('" +
                                                            myServerRelativeUrl + "')";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 20

//gavdcodebegin 28
static void SpCsRest_DeleteOneFolderInList()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myServerRelativeUrl = "/sites/[Site]/lists/[List]/RestFolderRenamed";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/web/GetFolderByServerRelativeUrl('" +
                                                            myServerRelativeUrl + "')";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 28

//gavdcodebegin 21
static void SpCsRest_CreateOneAttachment()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myFilePath = @"C:\Temporary\TestText.txt";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                   "/_api/lists/GetByTitle('TestList')" +
                        "/items(13)/AttachmentFiles/add(FileName='" + myFilePath + "')";

        object myPayloadObj = new{ };
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 21

//gavdcodebegin 22
static void SpCsRest_ReadAllAttachments()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                                "/_api/lists/GetByTitle('TestList')" +
                                                "/items(13)/AttachmentFiles";

        HttpClient myHttpClient = new HttpClient();
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
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 22

//gavdcodebegin 23
static void SpCsRest_DownloadOneAttachmentByFileName()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myFileName = "TestText.txt";
        string myFilesPath = @"C:\Temporary\";

        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                      "/_api/lists/GetByTitle('TestList')" +
                                      "/items(13)/AttachmentFiles('" + myFileName + "')" +
                                      "/$value";

        HttpClient myHttpClient = new HttpClient();
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Authorization", "Bearer " + myTokenWithAccPw.Item2);
        myHttpClient.DefaultRequestHeaders.Add(
                                    "Accept", "application/json"); // Output as JSON

        string resultStr = myHttpClient.GetAsync(myEndpoint).ContinueWith((myResponse) =>
        {
            return myResponse.Result.Content.ReadAsStringAsync().Result;
        }).Result;

        byte[] resultByte = Encoding.UTF8.GetBytes(resultStr);
        FileStream outputStream = new FileStream(myFilesPath + myFileName,
                            FileMode.OpenOrCreate | FileMode.Append,
                            FileAccess.Write, FileShare.None);
        outputStream.Write(resultByte, 0, resultByte.Length);
        outputStream.Flush(true);
        outputStream.Close();
    }
    else
    {
        Console.WriteLine(myTokenWithAccPw.Item2);
    }
}
//gavdcodeend 23

//gavdcodebegin 24
static void TestSpRSpCsRest_DeleteOneAttachmentByFileName()
{
    Tuple<string, string> myTokenWithAccPw = GetTokenWithAccPw();

    if (myTokenWithAccPw.Item1.ToLower() == "ok")
    {
        string myFileName = "TestText.txt";
        string myEndpoint = ConfigurationManager.AppSettings["SiteCollUrl"] +
                                        "/_api/lists/GetByTitle('TestList')" +
                                      "/items(13)/AttachmentFiles('" + myFileName + "')";

        object myPayloadObj = null;
        string myPayLoadJson = JsonConvert.SerializeObject(myPayloadObj);

        StringContent myStrContent = new StringContent(myPayLoadJson);
        myStrContent.Headers.ContentType = MediaTypeHeaderValue.Parse(
                                                    "application/json;odata=verbose");

        using (myStrContent)
        {
            HttpClient myHttpClient = new HttpClient();
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
//gavdcodeend 24


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsRest_CreateOneListItem();
//SpCsRest_UploadOneDocument();
//SpCsRest_DownloadOneDocumentt();
//SpCsRest_ReadAllListsItems();
//SpCsRest_ReadOneListsItem();
//SpCsRest_ReadAllLibraryDocs();
//SpCsRest_ReadOneLibraryDoc();
//SpCsRest_UpdateOneListItem();
//SpCsRest_UpdateOneLibraryDoc();
//SpCsRest_DeleteOneListItem();
//SpCsRest_DeleteOneLibraryDoc();
//SpCsRest_BreakSecurityInheritanceListItem();
//SpCsRest_ResetSecurityInheritanceListItem();
//SpCsRest_AddUserToSecurityRoleInListItem();
//SpCsRest_UpdateUserSecurityRoleInListItem();
//SpCsRest_DeleteUserFromSecurityRoleInListItem();
//SpCsRest_CreateOneFolderInLibrary();
//SpCsRest_CreateOneFolderInList();
//SpCsRest_ReadAllFoldersInLibrary();
//SpCsRest_ReadAllFoldersInList();
//SpCsRest_RenameOneFolderInLibrary();
//SpCsRest_RenameOneFolderInList();
//SpCsRest_DeleteOneFolderInLibrary();
//SpCsRest_DeleteOneFolderInList();
//SpCsRest_CreateOneAttachment();
//SpCsRest_ReadAllAttachments();
//SpCsRest_DownloadOneAttachmentByFileName();
//TestSpRSpCsRest_DeleteOneAttachmentByFileName();


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------



#nullable enable
