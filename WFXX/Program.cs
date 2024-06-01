using Newtonsoft.Json;
using System.Configuration;
using System.Net;
using System.Text;
using System.Web;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet 8.0 Console Application ***-----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 003
static void CsDotNet_GetTeamApp()
{
    string graphQuery =
    "https://graph.microsoft.com/v1.0/teams/dd1223a2-28a7-47d4-afc2-f42eae94f037";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        EndPoint = graphQuery,
        Method = HttpVerb.GET,
        Registration = TypeRegistration.Application
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsDotNet_GetTeamDel()
{
    string graphQuery =
    "https://graph.microsoft.com/v1.0/teams/bd71e9c8-edd3-4c61-8b1d-c4567769db5c";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"],
        EndPoint = graphQuery,
        Method = HttpVerb.GET,
        Registration = TypeRegistration.Delegation
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsDotNet_CreateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels";

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel 01 Application\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        EndPoint = graphQuery,
        Method = HttpVerb.POST,
        ContentType = "application/json",
        PostData = myBody,
        Registration = TypeRegistration.Application
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsDotNet_CreateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
                                "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels";

    string myBody = "{ " +
                        "\"displayName\": \"Graph Channel 02 Delegation\"," +
                        "\"description\": \"Channel created with Graph\"" +
                    " }";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"],
        EndPoint = graphQuery,
        Method = HttpVerb.POST,
        ContentType = "application/json",
        PostData = myBody,
        Registration = TypeRegistration.Delegation
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 006

static void CsDotNet_GetChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:eb21860817fb4fe1a774bef08091635d@thread.tacv2";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        EndPoint = graphQuery,
        Method = HttpVerb.GET,
        Registration = TypeRegistration.Application
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}

static void CsDotNet_GetChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:0da30c7628cb4b33923a49eb9f66141d@thread.tacv2";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"],
        EndPoint = graphQuery,
        Method = HttpVerb.GET,
        Registration = TypeRegistration.Delegation
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}

//gavdcodebegin 007
static void CsDotNet_UpdateChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:eb21860817fb4fe1a774bef08091635d@thread.tacv2";

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    List<HeaderConfig> myHeadersList = [];
    HeaderConfig myHeaderMat = new()
    {
        HeaderTitle = "IF-MATCH",
        HeaderValue = "*"
    };
    myHeadersList.Add(myHeaderMat);

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        EndPoint = graphQuery,
        Method = HttpVerb.PATCH,
        ContentType = "application/json",
        Headers = myHeadersList,
        PostData = myBody,
        Registration = TypeRegistration.Application
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsDotNet_UpdateChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:0da30c7628cb4b33923a49eb9f66141d@thread.tacv2";

    string myBody = "{ \"description\": \"Channel Description Updated\" }";

    List<HeaderConfig> myHeadersList = [];
    HeaderConfig myHeaderMat = new()
    {
        HeaderTitle = "IF-MATCH",
        HeaderValue = "*"
    };
    myHeadersList.Add(myHeaderMat);

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"],
        EndPoint = graphQuery,
        Method = HttpVerb.PATCH,
        ContentType = "application/json",
        Headers = myHeadersList,
        PostData = myBody,
        Registration = TypeRegistration.Delegation
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsDotNet_DeleteChannelApp()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:eb21860817fb4fe1a774bef08091635d@thread.tacv2";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        EndPoint = graphQuery,
        Method = HttpVerb.DELETE,
        Registration = TypeRegistration.Application
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsDotNet_DeleteChannelDel()
{
    string graphQuery = "https://graph.microsoft.com/v1.0/teams/" +
        "bd71e9c8-edd3-4c61-8b1d-c4567769db5c/channels/" +
        "19:0da30c7628cb4b33923a49eb9f66141d@thread.tacv2";

    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"],
        EndPoint = graphQuery,
        Method = HttpVerb.DELETE,
        Registration = TypeRegistration.Delegation
    };

    Tuple<string, string> resultText = myClient.SendGraphRequest();

    Console.WriteLine(resultText.Item1);
    Console.WriteLine(resultText.Item2);
}
//gavdcodeend 010

//gavdcodebegin 011
static AdAppToken CsDotNet_GetADTokenApplication()
{
    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithSecret"],
        ClientSecret = ConfigurationManager.AppSettings["ClientSecret"],
        TenantName = ConfigurationManager.AppSettings["TenantName"]
    };

    AdAppToken resultToken = myClient.GetAzureTokenApplication();

    return resultToken;
}
//gavdcodeend 011

//gavdcodebegin 012
static AdAppToken CsDotNet_GetADTokenDelegation()
{
    RestGraphClient myClient = new()
    {
        ClientID = ConfigurationManager.AppSettings["ClientIdWithAccPw"],
        TenantName = ConfigurationManager.AppSettings["TenantName"],
        UserName = ConfigurationManager.AppSettings["UserName"],
        UserPw = ConfigurationManager.AppSettings["UserPw"]
    };

    AdAppToken resultToken = myClient.GetAzureTokenDelegation();

    return resultToken;
}
//gavdcodeend 012

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 012 ***

//CsDotNet_GetTeamApp();
//CsDotNet_GetTeamDel();
//CsDotNet_CreateChannelApp();
//CsDotNet_CreateChannelDel();
//CsDotNet_GetChannelApp();   
//CsDotNet_GetChannelDel();   
//CsDotNet_UpdateChannelApp();
//CsDotNet_UpdateChannelDel();
//CsDotNet_DeleteChannelApp();
//CsDotNet_DeleteChannelDel();
//CSDotNet_AdAppToken myTokenApp = GetADTokenApplication(); Console.WriteLine(myTokenApp.access_token);
//CSDotNet_AdAppToken myTokenDel = GetADTokenDelegation(); Console.WriteLine(myTokenDel.access_token);

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
public class RestGraphClient
{
    public string ClientID { get; set; }
    public string ClientSecret { get; set; }
    public string TenantName { get; set; }
    public string EndPoint { get; set; }
    public HttpVerb Method { get; set; }
    public TypeRegistration Registration { get; set; }
    public string ContentType { get; set; }
    public string PostData { get; set; }
    public string UserName { get; set; }
    public string UserPw { get; set; }
    public List<HeaderConfig> Headers { get; set; }

    public RestGraphClient()
    {
    }

    public Tuple<string, string> SendGraphRequest()
    {
        AdAppToken adToken = new();
        if (Registration == TypeRegistration.Application)
            adToken = GetAzureTokenApplication();
        else if (Registration == TypeRegistration.Delegation)
            adToken = GetAzureTokenDelegation();

        if (adToken != null)
        {
            List<HeaderConfig> myHeadersList = [];
            HeaderConfig authorizationHeader = new()
            {
                HeaderTitle = "Authorization",
                HeaderValue = adToken.token_type + " " + adToken.access_token
            };
            myHeadersList.Add(authorizationHeader);
            Headers = myHeadersList;

            return SendGraphRequestInternal();
        }
        else
        {
            Tuple<string, string> tplReturn = new ("Error", string.Empty);
            return tplReturn;
        }
    }

    private Tuple<string, string> SendGraphRequestInternal()
    {
        HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(EndPoint);

        myRequest.Method = Method.ToString();
        myRequest.ContentLength = 0;
        myRequest.ContentType = ContentType;
        if (Headers != null)
        {
            foreach (HeaderConfig oneHeader in Headers)
            {
                myRequest.Headers.Add(oneHeader.HeaderTitle, oneHeader.HeaderValue);
            }
        }

        if (string.IsNullOrEmpty(PostData) == false)
        {
            byte[] bodyBytes = Encoding.GetEncoding("iso-8859-1").GetBytes(PostData);
            myRequest.ContentLength = bodyBytes.Length;

            using Stream writeStream = myRequest.GetRequestStream();
            writeStream.Write(bodyBytes, 0, bodyBytes.Length);
        }

        try
        {
            using HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            string responseValue = string.Empty;

            using (Stream responseStream = myResponse.GetResponseStream())
            {
                if (responseStream != null)
                    using (StreamReader myReader = new(responseStream))
                    {
                        responseValue = myReader.ReadToEnd();
                    }
            }

            Tuple<string, string> tplReturn = new (myResponse.StatusCode.ToString(), 
                                                   responseValue);
            return tplReturn;
        }
        catch (Exception ex)
        {
            Tuple<string, string> tplReturn = new ("Error", ex.ToString());
            return tplReturn;
        }
    }

    public AdAppToken GetAzureTokenApplication()
    {
        string LoginUrl = "https://login.microsoftonline.com";
        string ScopeUrl = "https://graph.microsoft.com/.default";

        string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";
        string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
            "grant_type=client_credentials&" +
            "client_id=" + ClientID + "&" +
            "client_secret=" + ClientSecret + "";

        RestGraphClient myClient = new()
        {
            EndPoint = myUri,
            Method = HttpVerb.POST,
            ContentType = "application/x-www-form-urlencoded",
            PostData = myBody
        };

        Tuple<string, string> tokenJSON = myClient.SendGraphRequestInternal();
        if (tokenJSON.Item1.Contains("Error") == false)
        {
            AdAppToken tokenObj =
                        JsonConvert.DeserializeObject<AdAppToken>(tokenJSON.Item2);
            return tokenObj;
        }

        return null;
    }

    public AdAppToken GetAzureTokenDelegation()
    {
        string LoginUrl = "https://login.microsoftonline.com";
        string ScopeUrl = "https://graph.microsoft.com/.default";

        string myUri = LoginUrl + "/" + TenantName + "/oauth2/v2.0/token";
        string myBody = "Scope=" + HttpUtility.UrlEncode(ScopeUrl) + "&" +
                        "grant_type=Password&" +
                        "client_id=" + ClientID + "&" +
                        "Username=" + UserName + "&" +
                        "Password=" + UserPw + "";

        RestGraphClient myClient = new()
        {
            EndPoint = myUri,
            Method = HttpVerb.POST,
            ContentType = "application/x-www-form-urlencoded",
            PostData = myBody
        };

        Tuple<string, string> tokenJSON = myClient.SendGraphRequestInternal();
        if (tokenJSON.Item1.Contains("Error") == false)
        {
            AdAppToken tokenObj =
                        JsonConvert.DeserializeObject<AdAppToken>(tokenJSON.Item2);
            return tokenObj;
        }

        return null;
    }
}
//gavdcodeend 001

//gavdcodebegin 002
public class AdAppToken
{
    public string token_type { get; set; }
    public string expires_in { get; set; }
    public string ext_expires_in { get; set; }
    public string expires_on { get; set; }
    public string not_before { get; set; }
    public string resource { get; set; }
    public string access_token { get; set; }
}

public enum HttpVerb
{
    GET,
    PATCH,
    POST,
    PUT,
    DELETE
}

public enum TypeRegistration
{
    Application,
    Delegation
}

public class HeaderConfig
{
    public string HeaderTitle { get; set; }
    public string HeaderValue { get; set; }
}
//gavdcodeend 002

#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
