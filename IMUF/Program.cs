using RestSharp;

//gavdcodebegin 001
// Legacy code
string myCard = CreateCard();
PostCard(myCard);
//gavdcodeend 001

//gavdcodebegin 002
static string CreateCard()  // Legacy code
{
    string picUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/" +
                        "b/b2/Microsoft-teams.jpg/120px-Microsoft-teams.jpg";

    return "{ " +
            "\"@type\": \"MessageCard\", " +
            "\"@context\": \"https://schema.org/extensions\", " +
            "\"summary\": \"Call from one user\", " +
            "\"themeColor\": \"0078D7\", " +
            "\"title\": \"Call opened: my WebHook is working\", " +
            "\"sections\": [ " +
            "    { " +
            "        \"activityTitle\": \"One user\", " +
            "        \"activitySubtitle\": \"" + DateTime.Now.ToString() + "\", " +
            "        \"activityImage\": \"" + picUrl + "\", " +
            "        \"facts\": [ " +
            "            { " +
            "                \"name\": \"Place:\", " +
            "                \"value\": \"Somewhere\" " +
            "            }, " +
            "            { " +
            "                \"name\": \"Call ID:\", " +
            "                \"value\": \"OneNumber\" " +
            "            } " +
            "        ], " +
            "        \"text\": \"There were no problems at all!\" " +
            "    } " +
            "], " +
            "\"potentialAction\": [ " +
            "    { " +
            "        \"@type\": \"ActionCard\", " +
            "        \"name\": \"Add a comment\", " +
            "        \"inputs\": [ " +
            "            { " +
            "                \"@type\": \"TextInput\", " +
            "                \"id\": \"comment\", " +
            "                \"title\": \"Enter your comment\", " +
            "                \"isMultiline\": true " +
            "            } " +
            "        ], " +
            "        \"actions\": [ " +
            "            { " +
            "                \"@type\": \"HttpPOST\", " +
            "                \"name\": \"OK\", " +
            "                \"target\": \"https://...\" " +
            "            } " +
            "        ] " +
            "    }, " +
            "    { " +
            "        \"@type\": \"HttpPOST\", " +
            "        \"name\": \"Close\", " +
            "        \"actions\": null, " +
            "        \"target\": \"https://...\" " +
            "    }, " +
            "    { " +
            "        \"@type\": \"OpenUri\", " +
            "        \"name\": \"Don't view it\", " +
            "        \"targets\": [ " +
            "            { " +
            "                \"os\": \"default\", " +
            "                \"uri\": \"https://...\" " +
            "            } " +
            "        ] " +
            "    } " +
            "] " +
        "}";
}
//gavdcodeend 002

//gavdcodebegin 003
static void PostCard(string theCard)  // Legacy code
{
    string WebhookUrl = "https://[domain].webhook.office.com/webhookb2/" +
        "28d184e1-60df-4bd2-9c48-c63b21943fbe@ade56059-89c0-4594-90c3-e4772a8168ca/" +
        "IncomingWebhook/657f5512286643e68938fd9b01476259/" +
        "acc28fcb-5261-47f8-960b-715d2f98a431";

    RestRequest myRequest = new();
    myRequest.AddHeader("content-type", "Application/Json");
    myRequest.AddJsonBody(theCard);

    RestClient myClient = new(WebhookUrl);
    RestResponse myResponse = myClient.ExecutePost(myRequest);

    if (myResponse.IsSuccessful == true)
        Console.WriteLine("WebHook sent successfully");
    else
        Console.WriteLine("Something went wrong");
}
//gavdcodeend 003
