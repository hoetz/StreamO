## Project Description ##

Listening to Exchange Streaming Notifications in a simple way.

Streaming notifications in Exchange (starting with 2010 SP1) combine the functionalities of push and pull notifications without having to set up a dedicated web service to receive the Exchange events. The client opens a long-standing call to EWS which is being used to receive real time notifications. StreamO aims to simplify Streaming Notification Subscriptions for many mailboxes and was created with Office 365 in mind.

##Requirements##

You need a service account for your Exchange environment with permissions to impersonate the targeted users. See http://msdn.microsoft.com/en-us/library/exchange/bb204095(v=exchg.140).aspx

Also, Autodiscover is a must.

##Sample usage##
            //Multiple long standing EWS calls might cause problems with the DefaultConnectionLimit
            ServicePointManager.DefaultConnectionLimit = 10;

            ExchangeCredentials cred = new WebCredentials("serviceaccount@yourdomain.com", "password");
            using (var listener = new StreamingListener(cred, (eventData) =>
                {
                    Console.WriteLine("Receiving events for {0}", eventData.Sender.ToString());
                    foreach (var e in eventData.Events)
                    {
                        Console.WriteLine(e.EventType.ToString());
                    }
                }))
            {
                //default: listen to NewMail events in the inbox
                listener.AddSubscription("florian.hoetzinger@yourdomain");
                
                //listen to Created events in the Contacts folder 
                listener.AddSubscription("john.doe@yourdomain.com",
                    new List<FolderId> { WellKnownFolderName.Contacts },
                    new List<EventType> { EventType.Created });

                Console.WriteLine("Listening...");
                Console.ReadLine();
            }
