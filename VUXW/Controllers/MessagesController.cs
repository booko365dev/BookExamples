using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Connector.Teams.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace VUXW.Controllers
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
        public async Task<HttpResponseMessage> Post([FromBody]Activity myActivity)
        {
            ComposeExtensionResponse myResponse = CreateCard(myActivity);
            return myResponse != null
                ? Request.CreateResponse<ComposeExtensionResponse>(myResponse)
                : new HttpResponseMessage(HttpStatusCode.OK);
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        private static ComposeExtensionResponse CreateCard(Activity myActivity)
        {
            ComposeExtensionResponse rtnResponse = null;

            dynamic activityValue = JObject.FromObject(myActivity.Value);

            string myFirst = activityValue.data.firstNumber;
            string mySecond = activityValue.data.secondNumber;

            int myAdd = int.Parse(myFirst) + int.Parse(mySecond);

            HeroCard myCard = new HeroCard
            {
                Title = "Add Card",
                Subtitle = "Adding " + myFirst + " + " + mySecond,
                Text = "The result is " + myAdd.ToString(),
                Images = new List<CardImage>(),
                Buttons = new List<CardAction>(),
            };
            myCard.Images.Add(new CardImage
            {
                Url = "http://wiki.opensemanticframework.org/images/0/0b/Add-72.png"
            });

            var myAttachs = new ComposeExtensionAttachment[1];
            myAttachs[0] = myCard.ToAttachment().ToComposeExtensionAttachment();

            rtnResponse = new ComposeExtensionResponse(
                                new ComposeExtensionResult("list", "result"));
            rtnResponse.ComposeExtension.Attachments = myAttachs.ToList();

            return rtnResponse;
        }
        //gavdcodeend 02

        // PUT: api/Messages/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE: api/Messages/5
        public void Delete(int id)
        {
        }
    }
}
