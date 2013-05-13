using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO.SampleApp
{
    internal static class App
    {
        private static void Main(string[] args)
        {
            ServicePointManager.DefaultConnectionLimit = 10;

            ExchangeCredentials cred = new WebCredentials("svcaccount@yourdomain.com", "password");
            using (var listener = new StreamingListener(cred, (x, y) =>
                {
                    foreach (var e in y.Events)
                    {
                        Console.WriteLine(e.EventType.ToString());
                    }
                }))
            {

                listener.AddSubscription(new MailAddress("florian.hoetzinger@yourdomain.com"),
                    new List<FolderId> { WellKnownFolderName.Contacts },
                    new List<EventType> { EventType.Created });

                listener.AddSubscription(new MailAddress("john.doe@yourdomain.com"),
                    new List<FolderId> { WellKnownFolderName.Contacts },
                    new List<EventType> { EventType.Created });

                listener.AddSubscription(new MailAddress("homer.simpson@yourdomain.com"),
                    new List<FolderId> { WellKnownFolderName.Contacts },
                    new List<EventType> { EventType.Created });

                Console.ReadLine();
            }
        }
    }
}