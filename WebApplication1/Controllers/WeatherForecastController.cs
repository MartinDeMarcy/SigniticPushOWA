using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.IO;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {

        [HttpGet]
        public ActionResult Get()
        {
            var json_string = "{ success: \"true\" }";
            string[] args = { "laurent@digital-optimizer.com", "C:/Users/Martin/source/repos/ConsoleApp1/ConsoleApp1/signature.html" };
            MainAsync(args).Wait();
            return Content(json_string, "application/json");
        }

        static void SetSigniture(ExchangeService service, string path)
        {
            service.TraceEnabled = true;
            Folder Root = Folder.Bind(service, WellKnownFolderName.Root);
            UserConfiguration OWAConfig = UserConfiguration.Bind(service, "OWA.UserOptions", Root.ParentFolderId, UserConfigurationProperties.All);

            // Open the stream and read it back.
            string signature = System.IO.File.ReadAllText(path);

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

            var app = ConfidentialClientApplicationBuilder.Create("a1f9cb88-0179-41bd-953e-291360018f02")
                .WithAuthority(AzureCloudInstance.AzurePublic, "0d9ba24c-e84b-47ff-a117-b91332d62420")
                .WithClientSecret("qMT1J2Ei5eu62_.l~ADQ~dTPamT3JepO-H").Build();
            try
            {
                // Make the interactive token request
                AuthenticationResult result = await app.AcquireTokenForClient(ewsScopes)
                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService
                {
                    Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                    Credentials = new OAuthCredentials(result.AccessToken),

                    //Impersonate the mailbox you'd like to access.
                    ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, args[0])
                };
                SetSigniture(ewsClient, args[1]);

            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }
        }
    }
}

