using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;

namespace JCIW
{
    class Program
    {
        // Note: Remove one of the two Main routines to run the program
        //gavdcodebegin 02
        static void Main(string[] args)
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["exUserName"],
                                ConfigurationManager.AppSettings["exUserPw"]);

            CallEWSTest(myExService);
        }
        //gavdcodeend 02

        //gavdcodebegin 04
        static void Main(string[] args)
        {
            ExchangeService myExService = ConnectOA(
                                ConfigurationManager.AppSettings["exAppId"],
                                ConfigurationManager.AppSettings["exTenantId"]).
                                                            GetAwaiter().GetResult();

            CallEWSTest(myExService);
        }
        //gavdcodeend 04

        //gavdcodebegin 01
        static ExchangeService ConnectBA(string userEmail, string userPW)
        {
            ExchangeService exService = new ExchangeService
            {
                Credentials = new WebCredentials(userEmail, userPW)
            };

            //exService.TraceEnabled = true;
            //exService.TraceFlags = TraceFlags.All;

            exService.AutodiscoverUrl(userEmail, RedirectionUrlValidationCallback);
            //Console.WriteLine(exService.Url);

            return exService;
        }

        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool validationResult = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                validationResult = true;
            }

            return validationResult;
        }
        //gavdcodeend 01

        //gavdcodebegin 03
        static async System.Threading.Tasks.Task<ExchangeService> ConnectOA(
                                                            string AppId, string TenId)
        {
            ExchangeService exService = new ExchangeService();

            PublicClientApplicationOptions pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = AppId,
                TenantId = TenId
            };

            IPublicClientApplication pcaBuilder = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            string[] exScope = new string[] {
                            "https://outlook.office.com/EWS.AccessAsUser.All" };

            AuthenticationResult authToken = await
                              pcaBuilder.AcquireTokenInteractive(exScope).ExecuteAsync();

            exService.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            exService.Credentials = new OAuthCredentials(authToken.AccessToken);

            return await System.Threading.Tasks.Task.FromResult(exService);
        }
        //gavdcodeend 03

        static void CallEWSTest(ExchangeService ExchService)
        {
            FindFoldersResults allFolders = ExchService.FindFolders(WellKnownFolderName.MsgFolderRoot,
                                                                    new FolderView(10));
            foreach (Folder oneFolder in allFolders)
            {
                Console.WriteLine(oneFolder.DisplayName);
            }
        }
    }
}
