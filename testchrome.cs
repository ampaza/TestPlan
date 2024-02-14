
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI; 
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;

namespace GmailAPISample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Authenticate using credentials file
            UserCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { GmailService.Scope.GmailReadonly },
                    "user",
                    System.Threading.CancellationToken.None,
                    new Google.Apis.Util.Store.FileDataStore("token.json", true)).Result;
            }

            // Create Gmail service
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail API Sample",
            });

            
            // Get user's email messages
            ListMessages(service, "me");
        }

        static void ListMessages(GmailService service, string userId)
        {
            // Define parameters of request
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.LabelIds = "INBOX"; // Limit to emails in the inbox
            request.Q = "category:primary"; // Only emails in the "Primary" tab
            request.MaxResults = 30; // Limit to the first email

            try
            {
                // List messages
                IList<Message> messages = request.Execute().Messages;
                if (messages != null && messages.Count > 0)
                {
                    Console.WriteLine("Total number of emails in the Primary tab: " + messages.Count);

                    // Get the first message
                    var firstMessage = messages[0];

                    // Get the message details
                    var email = service.Users.Messages.Get(userId, firstMessage.Id).Execute();
                    var headers = email.Payload.Headers;
                    string sender = "";
                    string subject = "";

                    // Find sender and subject headers
                    foreach (var header in headers)
                    {
                        if (header.Name == "From")
                        {
                            sender = header.Value;
                        }
                        else if (header.Name == "Subject")
                        {
                            subject = header.Value;
                        }
                    }

                    // Print sender and subject of the first email
                    Console.WriteLine("Sender: {0}\nSubject: {1}", sender, subject);
                }
                else
                {
                    Console.WriteLine("No emails found in the Primary tab.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}


