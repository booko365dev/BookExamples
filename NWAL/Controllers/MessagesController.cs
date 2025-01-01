using Microsoft.AspNet.WebHooks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace NWAL.Controllers
{
    //gavdcodebegin 001
    public class MessagesController : WebHookHandler  // Legacy code
    {
        //gavdcodeend 001

        //gavdcodebegin 002
        public MessagesController()  // Legacy code
        {
            this.Receiver = GenericJsonWebHookReceiver.ReceiverName;
        }

        public override Task ExecuteAsync(string TheReceiver,
                                     WebHookHandlerContext TheContext) // Legacy code
        {
            string myBodyFormatted = TheContext.Data.ToString();
            dynamic myBodyObj = JsonConvert.DeserializeObject(myBodyFormatted);

            bool myValidationResult = ValidationIsOk(TheContext);

            if (myValidationResult == true)
            {
                SendMessageBack(myBodyObj, myValidationResult, TheContext);
            }

            return Task.FromResult(true);
        }
        //gavdcodeend 002

        //gavdcodebegin 004
        static void SendMessageBack(dynamic BodyObj, bool ValidationResult,
                                    WebHookHandlerContext TheContext)  // Legacy code
        {
            string myFrom = BodyObj.from.name;
            string myText = BodyObj.text;

            string jsonMessage = "{ \"type\": \"message\", \"text\": \"Request '" +
                myText.Trim() + "' from '" + myFrom + "' accepted. Validation is " +
                ValidationResult.ToString() + "\" }";
            TheContext.Response = TheContext.Request.CreateResponse();
            TheContext.Response.Content = new StringContent(jsonMessage);
        }
        //gavdcodeend 004

        //gavdcodebegin 003
        static bool ValidationIsOk(WebHookHandlerContext TheContext)  // Legacy code
        {
            string myHMAC_Calculated = string.Empty;
            string signingKey = "BVK3o93NuB6y3BA5J5k+mBR9n+mG+qGDWbtcxHcrUYg=";

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

            bool rtnBool = myHMAC_Authorization.Equals(myHMAC_Calculated);

            return rtnBool;
        }
        //gavdcodeend 003
    }
}
