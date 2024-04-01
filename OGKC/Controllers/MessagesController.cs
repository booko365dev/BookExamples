using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace OGKC.Controllers
{
    public class MessagesController : ApiController
    {
        // GET: api/Messages
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET: api/Messages/5
        public string Get(int id)
        {
            return "value";
        }

        //gavdcodebegin 001
        // ATTENTION: Replaced by QEXM
        [HttpPost]
        [BotAuthentication]
        public async Task<HttpResponseMessage> Post([FromBody] Activity myActivity)
        {
            using (ConnectorClient myConnector = new ConnectorClient(new Uri(myActivity.ServiceUrl),
                ConfigurationManager.AppSettings[
                                      MicrosoftAppCredentials.MicrosoftAppIdKey],
                ConfigurationManager.AppSettings[
                                      MicrosoftAppCredentials.MicrosoftAppPasswordKey]))
            {
                string[] wordsInText = myActivity.GetTextWithoutMentions().Split(' ');
                string afterAbout = string.Empty;
                for (int myCounter = 0; myCounter < wordsInText.Length; myCounter++)
                {
                    if (wordsInText[myCounter].Trim().ToLower() == "about")
                    {
                        afterAbout = wordsInText[myCounter + 1];
                        break;
                    }
                }

                string myResult = string.Empty;
                if (string.IsNullOrEmpty(afterAbout) == false)
                {
                    myResult = await GetWikipediaSnippet(afterAbout);
                }
                else
                {
                    myResult = "Nothing found about '" + afterAbout + "'";
                }

                Activity myReply = myActivity.CreateReply("About '" +
                                                    afterAbout + "': " + myResult);
                await myConnector.Conversations.ReplyToActivityWithRetriesAsync(myReply);

                return new HttpResponseMessage(HttpStatusCode.Accepted);
            }
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        // ATTENTION: Replaced by QEXM
        private async Task<string> GetWikipediaSnippet(string WordToQuery)
        {
            string strReturn = string.Empty;

            HttpClient client = new HttpClient();

            string wikiUrl = "https://en.wikipedia.org/w/api.php?" +
                "action=query&list=search&srsearch=" + WordToQuery +
                "&utf8=&format=json";

            client.BaseAddress = new Uri(wikiUrl);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

            Wikipedia myResult = null;
            HttpResponseMessage response = await client.GetAsync(client.BaseAddress);
            if (response.IsSuccessStatusCode)
            {
                myResult = await response.Content.ReadAsAsync<Wikipedia>();
            }

            strReturn = myResult.query.search[0].snippet;
            return strReturn;
        }
        //gavdcodeend 002

        // PUT: api/Messages/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/Messages/5
        public void Delete(int id)
        {
        }
    }

    //gavdcodebegin 003
    // ATTENTION: Replaced by QEXM
    public class Wikipedia
    {
        public string batchcomplete { get; set; }
        public Continue _continue { get; set; }
        public Query query { get; set; }
    }

    public class Continue
    {
        public int sroffset { get; set; }
        public string _continue { get; set; }
    }

    public class Query
    {
        public Searchinfo searchinfo { get; set; }
        public Search[] search { get; set; }
    }

    public class Searchinfo
    {
        public int totalhits { get; set; }
    }

    public class Search
    {
        public int ns { get; set; }
        public string title { get; set; }
        public int pageid { get; set; }
        public int size { get; set; }
        public int wordcount { get; set; }
        public string snippet { get; set; }
        public DateTime timestamp { get; set; }
    }
    //gavdcodeend 003
}
