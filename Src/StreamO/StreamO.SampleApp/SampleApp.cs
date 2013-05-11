using System;
using System.Collections.Generic;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO.SampleApp
{
    internal static class App
    {
        private static void Main(string[] args)
        {
            ExchangeCredentials cred = new WebCredentials("your@serviceaccount.com", "password");
            var listener = new StreamingListener(cred, ExchangeVersion.Exchange2010_SP2, (x, y) =>
                {
                    foreach (var e in y.Events)
                    {
                        Console.WriteLine(e.EventType.ToString());
                    }
                });

            listener.AddSubscription(new MailAddress("florian.hoetzinger@gab-net.com"),
                new List<FolderId> { WellKnownFolderName.Contacts },
                new List<EventType> { EventType.Created });

            Console.ReadLine();
        }
    }
}