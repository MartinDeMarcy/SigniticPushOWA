using System;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.IO;

namespace OWAUtility
{
    public class OWA
    {
        static void Main(string[] args)
        {

            MainAsync(args).Wait();

        }

        static void SetSigniture(ExchangeService service, string path)
        {
            service.TraceEnabled = true;
            Folder Root = Folder.Bind(service, WellKnownFolderName.Root);
            UserConfiguration OWAConfig = UserConfiguration.Bind(service, "OWA.UserOptions", Root.ParentFolderId, UserConfigurationProperties.All);

            // Open the stream and read it back.
            string signature = File.ReadAllText(path);

            if (OWAConfig.Dictionary.ContainsKey("signaturehtml"))
            {
                OWAConfig.Dictionary["signaturehtml"] = signature;
            }
            else
            {
                OWAConfig.Dictionary.Add("signaturehtml", signature);
            }

            OWAConfig.Update();
        }

        public static async System.Threading.Tasks.Task MainAsync(string[] args)
        {
            // Configure the MSAL client to get tokens
            var ewsScopes = new string[] { "https://outlook.office.com/.default" };

            var app = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["appId"])
                .WithAuthority(AzureCloudInstance.AzurePublic, ConfigurationManager.AppSettings["tenantId"])
                .WithClientSecret(ConfigurationManager.AppSettings["clientSecret"])
                .Build();

            AuthenticationResult result = null;

            try
            {
                // Make the interactive token request
                result = await app.AcquireTokenForClient(ewsScopes)
                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(result.AccessToken);

                //Impersonate the mailbox you'd like to access.
                ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, args[0]);
                SetSigniture(ewsClient, args[1]);

            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex.ToString()}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }
        }
    }
}
