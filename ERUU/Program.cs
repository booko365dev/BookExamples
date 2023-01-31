using RestSharp;
using System;

namespace ERUU
{
    class Program
    {
        //gavdcodebegin 001
        static void Main(string[] args)
        {
            string myCard = CreateCard();
            PostCard(myCard);
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        static string CreateCard()
        {
            string picUrl = "https://upload.wikimedia.org/wikipedia/commons/thumb/" +
                                "b/b2/Microsoft-teams.jpg/120px-Microsoft-teams.jpg";

            return "{ " +
                    "\"@type\": \"MessageCard\", " +
                    "\"@context\": \"https://schema.org/extensions\", " +
                    "\"summary\": \"Call from one user\", " +
                    "\"themeColor\": \"0078D7\", " +
                    "\"title\": \"Call opened: something is not working\", " +
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
                    "        \"text\": \"There is a problem somewhere\" " +
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
                    "                \"target\": \"http://...\" " +
                    "            } " +
                    "        ] " +
                    "    }, " +
                    "    { " +
                    "        \"@type\": \"HttpPOST\", " +
                    "        \"name\": \"Close\", " +
                    "        \"actions\": null, " +
                    "        \"target\": \"http://...\" " +
                    "    }, " +
                    "    { " +
                    "        \"@type\": \"OpenUri\", " +
                    "        \"name\": \"Don't view it\", " +
                    "        \"targets\": [ " +
                    "            { " +
                    "                \"os\": \"default\", " +
                    "                \"uri\": \"http://...\" " +
                    "            } " +
                    "        ] " +
                    "    } " +
                    "] " +
                "}";
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        static void PostCard(string theCard)
        {
            string WebhookUrl = "https://outlook.office.com/webhook/3a0c86a6-4bb9-" +
                "4846-b712-fea17c4542e8@03d561bf-4472-41e0-b2d6-ee506471e9d0/" +
                "IncomingWebhook/1ff55c8a6073460d94867869922b09d8/092b1237-a428-" +
                "45a7-b76b-310fdd6e7246";

            RestRequest myRequest = new RestRequest(Method.POST);
            myRequest.AddHeader("content-type", "Application/Json");
            myRequest.AddJsonBody(theCard);

            RestClient myClient = new RestClient(WebhookUrl);
            myClient.Execute(myRequest);
        }
        //gavdcodeend 003
    }
}
