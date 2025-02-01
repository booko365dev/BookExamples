using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ORGK
{
    public class SpWebhookListenerFunction
    {
        private readonly ILogger<SpWebhookListenerFunction> _logger;

        public SpWebhookListenerFunction(ILogger<SpWebhookListenerFunction> logger)
        {
            _logger = logger;
        }

        //gavdcodebegin 001
        [Function("SharePointWebhookReceiver")]
        public IActionResult Run(
                    [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequest req)
        {
            string currentDateTime = DateTime.Now.ToString("yyyy-dd-MM;HH-mm-ss");
            _logger.LogInformation(
                "------------ Webhook Listener - {currentDateTime}", currentDateTime);

            string myValidationToken = req.Query["validationtoken"]!;

            // Respond to validation requests from SharePoint
            if (!string.IsNullOrEmpty(myValidationToken))
            {
                _logger.LogInformation(
                    "Received validation token: {myValidationToken}", myValidationToken);

                ContentResult myResult =
                    CreateContentResult(myValidationToken, HttpStatusCodeEnum.OK);
                _logger.LogInformation(
                    "Sent back validation token: {ValidationToken}", myResult.Content);

                return myResult;
            }

            // Process webhook notifications
            SPWebhookNotification? myNotification = null;

            string myRequestBodyContent = new
                        StreamReader(req.Body).ReadToEndAsync().Result;
            _logger.LogInformation(
                "Body received - {RequestBodyContent}", myRequestBodyContent);

            SPWebhookContent? objNotification =
                JsonConvert.DeserializeObject<SPWebhookContent>(myRequestBodyContent);
            if (objNotification?.Value != null && objNotification.Value.Count > 0)
            {
                myNotification = objNotification.Value[0];
            }

            if (myNotification != null)
            {
                Task.Factory.StartNew(() =>
                {
                    // Handle the myNotification here

                    _logger.LogInformation(
                        "- Resource: {Resource}", myNotification.Resource);
                    _logger.LogInformation(
                        "- ClientState: {ClientState}", myNotification.ClientState);
                    _logger.LogInformation(
                        "- SubscrId: {SubscriptionId}", myNotification.SubscriptionId);
                    _logger.LogInformation(
                        "- TenantId: {TenantId}", myNotification.TenantId);
                    _logger.LogInformation(
                        "- SiteUrl: {SiteUrl}", myNotification.SiteUrl);
                    _logger.LogInformation(
                        "- WebId: {WebId}", myNotification.WebId);
                    _logger.LogInformation(
                        "- ExpDateTime: {ExpDatTim}", myNotification.ExpirationDateTime);
                });

                return CreateContentResult(string.Empty, HttpStatusCodeEnum.OK);
            }

            return CreateContentResult(string.Empty, HttpStatusCodeEnum.BadRequest);
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        private static ContentResult CreateContentResult(string strContent,
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
