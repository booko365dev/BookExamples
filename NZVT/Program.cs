using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Configuration;
using System.Security;

namespace NZVT
{
    class Program
    {
        static void Main(string[] args)
        {
            string tenantId = ConfigurationManager.AppSettings["TenantName"];
            string clientIdApp = ConfigurationManager.AppSettings["ClientIdApp"];
            string clientSecretApp = ConfigurationManager.AppSettings["ClientSecretApp"];
            string clientIdDel = ConfigurationManager.AppSettings["ClientIdDel"];

            //GetGrQueryApp(tenantId, clientIdApp, clientSecretApp);
            //GetGrQueryDel(tenantId, clientIdDel);
            //GetTokenApp(tenantId, clientIdApp, clientSecretApp);
            //GetTokenDel(tenantId, clientIdDel);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 03
        static void GetGrQueryApp(string TenantId, string ClientId, string ClientSecret)
        {
            GraphServiceClient graphClient = GetGraphClientApp(
                                                    TenantId, ClientId, ClientSecret);

            User myMe = graphClient.Users[ConfigurationManager.AppSettings["UserName"]]
                .Request()
                .GetAsync().Result;
            Console.WriteLine(myMe.DisplayName);

            var myMessages = graphClient.Users[ConfigurationManager.AppSettings["UserName"]]
                .Messages.Request()
                .GetAsync().Result;
            Console.WriteLine(myMessages.Count.ToString());
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void GetGrQueryDel(string TenantId, string ClientId)
        {
            GraphServiceClient graphClient = GetGraphClientDel(TenantId, ClientId);

            var securePassword = new SecureString();
            foreach (var chr in ConfigurationManager.AppSettings["UserPw"]) 
                                                    { securePassword.AppendChar(chr); }

            User myMe = graphClient.Me.Request()
                .WithUsernamePassword(ConfigurationManager.AppSettings
                                                        ["UserName"], securePassword)
                .GetAsync().Result;
            Console.WriteLine(myMe.DisplayName);

            var myMessages = graphClient.Me.Messages.Request()
                .WithUsernamePassword(ConfigurationManager.AppSettings
                                                        ["UserName"], securePassword)
                .GetAsync().Result;
            Console.WriteLine(myMessages.Count.ToString());
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void GetTokenApp(string TenantId, string ClientId, string ClientSecret)
        {
            IConfidentialClientApplication clientApplication = 
                ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithTenantId(TenantId)
                    .WithClientSecret(ClientSecret)
                    .Build();

            string[] myScopes = new string[] { "https://graph.microsoft.com/.default" };

            var myToken = clientApplication
                    .AcquireTokenForClient(myScopes)
                    .ExecuteAsync().Result;

            Console.WriteLine("Token value - " + myToken.AccessToken);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void GetTokenDel(string TenantId, string ClientId)
        {
            IPublicClientApplication clientApplication = PublicClientApplicationBuilder
                    .Create(ClientId)
                    .WithTenantId(TenantId)
                    .Build();

            string[] myScopes = new string[] { "https://graph.microsoft.com/.default" };

            var securePassword = new SecureString();
            foreach (var chr in ConfigurationManager.AppSettings["UserPw"]) 
                                                { securePassword.AppendChar(chr); }

            var myToken = clientApplication
                .AcquireTokenByUsernamePassword(myScopes,
                                        ConfigurationManager.AppSettings
                                                        ["UserName"], securePassword)
                .ExecuteAsync().Result;

            Console.WriteLine("Token for - " + myToken.Account.Username);
            Console.WriteLine("Token value - " + myToken.AccessToken);
        }
        //gavdcodeend 06

//---------------------------------------------------------------------------------------

        //gavdcodebegin 01
        static GraphServiceClient GetGraphClientApp(string TenantId, string ClientId, 
                                                                    string ClientSecret)
        {
            IConfidentialClientApplication clientApplication = 
                ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithTenantId(TenantId)
                    .WithClientSecret(ClientSecret)
                    .Build();

            ClientCredentialProvider authenticationProvider = 
                                    new ClientCredentialProvider(clientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);

            return graphClient;
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static GraphServiceClient GetGraphClientDel(string TenantId, string ClientId)
        {
            IPublicClientApplication clientApplication = PublicClientApplicationBuilder
                .Create(ClientId)
                .WithTenantId(TenantId)
                //.WithRedirectUri("http://localhost") // Only if redirect in the App Reg.
                .Build();

            UsernamePasswordProvider authenticationProvider = 
                                    new UsernamePasswordProvider(clientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);

            return graphClient;
        }
        //gavdcodeend 02
    }
}
