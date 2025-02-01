using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using System.Web;

namespace DYCH.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SpWebhookListenerController : ControllerBase
    {
        // GET: api/<SpWebhookListenerController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<SpWebhookListenerController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        //gavdcodebegin 001
        // POST api/<SpWebhookListenerController>
        [HttpPost]
        public IActionResult Post()
        {
            string currentDateTime = DateTime.Now.ToString("yyyy-dd-MM;HH-mm-ss");
            Console.WriteLine("----------------------------- POST " + currentDateTime);

            HttpResponseMessage myHttpResponse = new(HttpStatusCode.BadRequest);
            IEnumerable<string> myClientStateHeader = [];
            string myWebhookClientState = "guitacaClientState";

            if (Request.Headers.TryGetValue("ClientState",
                                                out var myClientStateHeaderValues))
            {
                string oneClientStateHeaderValue =
                    myClientStateHeaderValues.FirstOrDefault() ?? string.Empty;

                if (!string.IsNullOrEmpty(oneClientStateHeaderValue) &&
                        oneClientStateHeaderValue.Equals(myWebhookClientState))
                {
                    Console.WriteLine("Received client state: " +
                        oneClientStateHeaderValue);

                    var myQueryStringParams =
                            HttpUtility.ParseQueryString(Request.QueryString.Value!);

                    if (myQueryStringParams.AllKeys.Contains("validationtoken"))
                    {
                        // Subscription request received
                        string myValidationToken = myQueryStringParams
                                        .GetValues("validationtoken")![0].ToString();
                        Console.WriteLine("Received validation token: " +
                                        myValidationToken);

                        ContentResult myResult = CreateContentResult(myValidationToken,
                                        HttpStatusCodeEnum.OK);
                        Console.WriteLine("Sent back validation token: " +
                                        myResult.Content!.ToString());

                        return myResult;
                    }
                    else
                    {
                        // Notification received
                        using StreamReader reader = new(Request.Body, Encoding.UTF8);
                        string myRequestBodyContent = reader.ReadToEndAsync().Result;
                        Console.WriteLine("Body received - " + myRequestBodyContent);

                        if (!string.IsNullOrEmpty(myRequestBodyContent))
                        {
                            SPWebhookNotification? myNotification = null;

                            try
                            {
                                SPWebhookContent? objNotification =
                                    JsonConvert.DeserializeObject<SPWebhookContent>(
                                            myRequestBodyContent);
                                if (objNotification?.Value != null &&
                                            objNotification.Value.Count > 0)
                                {
                                    myNotification = objNotification.Value[0];
                                }
                            }
                            catch (Newtonsoft.Json.JsonException ex)
                            {
                                Console.WriteLine("JSON deserialization error: " +
                                            ex.InnerException);
                                return CreateContentResult(string.Empty,
                                            HttpStatusCodeEnum.BadRequest);
                            }

                            if (myNotification != null)
                            {
                                Task.Factory.StartNew(() =>
                                {
                                    // Handle the myNotification here

                                    Console.WriteLine("- Notification Resource: " +
                                                myNotification.Resource);
                                    Console.WriteLine("- Notification ClientState: " +
                                                myNotification.ClientState);
                                    Console.WriteLine("- Notification SubscriptionId: " +
                                                myNotification.SubscriptionId);
                                    Console.WriteLine("- Notification TenantId: " +
                                                myNotification.TenantId);
                                    Console.WriteLine("- Notification SiteUrl: " +
                                                myNotification.SiteUrl);
                                    Console.WriteLine("- Notification WebId: " +
                                                myNotification.WebId);
                                    Console.WriteLine("- Notification ExpDateTime: " +
                                                myNotification.ExpirationDateTime);
                                });

                                return CreateContentResult(string.Empty,
                                                HttpStatusCodeEnum.OK);
                            }
                        }
                    }
                }
                else
                {
                    return CreateContentResult(string.Empty, HttpStatusCodeEnum.Forbiden);
                }
            }

            return CreateContentResult(string.Empty, HttpStatusCodeEnum.BadRequest);
        }
        //gavdcodeend 001

        // PUT api/<SpWebhookListenerController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<SpWebhookListenerController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }

        //gavdcodebegin 002
        private ContentResult CreateContentResult(string strContent,
                                                  HttpStatusCodeEnum statusCode)
        {
            return new ContentResult
            {
                Content = strContent,
                ContentType = "text/plain",
                StatusCode = (int)statusCode
            };
        }
        //gavdcodeend 002
    }

    //gavdcodebegin 003
    public class SPWebhookNotification
    {
        public string? Resource { get; set; }
        public string? SubscriptionId { get; set; }
        public string? TenantId { get; set; }
        public string? SiteUrl { get; set; }
        public string? WebId { get; set; }
        public string? ClientState { get; set; }
        public DateTime ExpirationDateTime { get; set; }
    }

    public class SPWebhookContent
    {
        public List<SPWebhookNotification>? Value { get; set; }
    }

    public enum HttpStatusCodeEnum
    {
        OK = 200,  //HttpStatusCode.OK
        BadRequest = 400,  //HttpStatusCode.BadRequest
        Forbiden = 403  //HttpStatusCode.Forbiden
    }
    //gavdcodeend 003
}
