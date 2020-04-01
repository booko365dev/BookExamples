using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;

namespace JHCW.Controllers
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

        //gavdcodebegin 01
        [HttpPost]
        public async Task<HttpResponseMessage> Post([FromBody] Activity myActivity)
        {
            ComposeExtensionResponse myResponse = await CreateCard(myActivity);

            return myResponse != null
                ? Request.CreateResponse<ComposeExtensionResponse>(myResponse)
                : new HttpResponseMessage(HttpStatusCode.OK);
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        private async Task<ComposeExtensionResponse> CreateCard(Activity myActivity)
        {
            ComposeExtensionResponse rtnResponse = null;

            ComposeExtensionQuery queryData = myActivity.GetComposeExtensionQueryData();
            string myQuery = string.Empty;
            ComposeExtensionParameter queryParam = queryData.Parameters?.FirstOrDefault(
                                            par => par.Name == "wikipediaQuery");
            if (queryParam != null)
            {
                myQuery = queryParam.Value.ToString();
            }

            string myText = string.Empty;
            if (string.IsNullOrEmpty(myQuery) == true)
                myText = "Get a query to create the card";
            else
                myText = myQuery;

            string wikiResult = await GetWikipediaSnippet(myText);

            HeroCard myCard = new HeroCard
            {
                Title = "Wikipedia Card",
                Subtitle = "Searching for: " + myText,
                Text = wikiResult,
                Images = new List<CardImage>(),
                Buttons = new List<CardAction>(),
            };
            myCard.Images.Add(new CardImage
            {
                Url = "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/" +
                            "Orange_Icon_Pale_Wiki.png/120px-Orange_Icon_Pale_Wiki.png"
            });

            ComposeExtensionAttachment[] myAttachs = new ComposeExtensionAttachment[1];
            myAttachs[0] = myCard.ToAttachment().ToComposeExtensionAttachment();

            rtnResponse = new ComposeExtensionResponse(
                                    new ComposeExtensionResult("list", "result"));
            rtnResponse.ComposeExtension.Attachments = myAttachs.ToList();

            return rtnResponse;
        }
        //gavdcodeend 02

        //gavdcodebegin 03
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
        //gavdcodeend 03

        // DELETE: api/Messages/5
        public void Delete(int id)
        {
        }
    }

    //gavdcodebegin 04
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
    //gavdcodeend 04
}
