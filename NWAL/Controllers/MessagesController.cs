using Microsoft.AspNet.WebHooks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace NWAL.Controllers
{
    //gavdcodebegin 01
    public class MessagesController : WebHookHandler
    {
        //gavdcodeend 01

        //gavdcodebegin 02
        public MessagesController()
        {
            this.Receiver = GenericJsonWebHookReceiver.ReceiverName;
        }

        public override Task ExecuteAsync(string TheReceiver,
                                          WebHookHandlerContext TheContext)
        {
            string myAction = TheContext.Actions.First();
            string myBodyFormated = TheContext.Data.ToString();
            dynamic myBodyObj = JsonConvert.DeserializeObject(myBodyFormated);

            bool myValidationResult = ValidationIsOk(TheContext);

            if (myValidationResult == true)
            {
                SendMessageBack(myBodyObj, myValidationResult, TheContext);
            }

            return Task.FromResult(true);
        }
        //gavdcodeend 02

        //gavdcodebegin 04
        static void SendMessageBack(dynamic BodyObj, bool ValidationResult,
                                    WebHookHandlerContext TheContext)
        {
            string myFrom = BodyObj.from.name;
            string myText = BodyObj.text;

            string jsonMessage = "{ \"type\": \"message\", \"text\": \"Request '" +
                myText.Trim() + "' from '" + myFrom + "' accepted. Validation is " +
                ValidationResult.ToString() + "\" }";
            TheContext.Response = TheContext.Request.CreateResponse();
            TheContext.Response.Content = new StringContent(jsonMessage);
        }
        //gavdcodeend 04

        //gavdcodebegin 03
        static bool ValidationIsOk(WebHookHandlerContext TheContext)
        {
            bool rtnBool = false;
            string myHMAC_Calculated = string.Empty;
            string signingKey = "LtHurznxPxnvpew5GRMNooIuF6kpfzjxwYpXnJVEONo=";

            var myHMACFromAuthorization = TheContext.Request.Headers.Authorization;
            string myHMAC_Authorization = myHMACFromAuthorization.Parameter;

            JObject myBodyMinimizedObj = TheContext.GetDataOrDefault<JObject>();
            string myDataSerialized = JsonConvert.SerializeObject(myBodyMinimizedObj);
            byte[] myDataBytes = Encoding.UTF8.GetBytes(myDataSerialized);

            byte[] signingKeyBytes = Convert.FromBase64String(signingKey);
            using (HMACSHA256 myHMAC_SHA256 = new HMACSHA256(signingKeyBytes))
            {
                byte[] myDataHashBytes = myHMAC_SHA256.ComputeHash(myDataBytes);
                myHMAC_Calculated = Convert.ToBase64String(myDataHashBytes);
            }

            rtnBool = myHMAC_Authorization.Equals(myHMAC_Calculated);

            return rtnBool;
        }
        //gavdcodeend 03
    }
}
