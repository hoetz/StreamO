using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO
{
    public class StreamingListener
    {
        private readonly ExchangeCredentials _credentials;
        private readonly ExchangeVersion _exchangeVersion;
        private readonly IList<StreamingSubscriptionCollection> _subscriptionCollections = new List<StreamingSubscriptionCollection>();
        private readonly Action<object, NotificationEventArgs> _onNotificationEvent;

        /// <summary>
        /// Manages Streaming Notifications for Exchange Users. Automatically assigns subscriptions to adequate CAS connections.
        /// </summary>
        /// <param name="credentials">Credentials with permission to impersonate the user mailboxes for all the subscriptions this instance will manage.</param>
        /// <param name="exchangeVersion">The version of the target Exchange server. Must be 2010 SP1 or higher</param>
        /// <param name="onNotificationEvent">The Action to invoke when Notifications arrive</param>
        public StreamingListener(ExchangeCredentials credentials, ExchangeVersion exchangeVersion, Action<object, NotificationEventArgs> onNotificationEvent)
        {
            _onNotificationEvent = onNotificationEvent;
            if ((int)exchangeVersion < 2)
                throw new ArgumentException("ExchangeVersion must be 2010 SP1 or higher");
            _exchangeVersion = exchangeVersion;
            this._credentials = credentials;
        }

        /// <summary>
        /// Manages Streaming Notifications for Exchange Users. Automatically assigns subscriptions to adequate CAS connections. Exchange Version is assumed to be 2010_SP1
        /// </summary>
        /// <param name="credentials">Credentials with permission to impersonate the user mailboxes for all the subscriptions this instance will manage.</param>
        /// <param name="onNotificationEvent">The Action to invoke when Notifications arrive</param>
        public StreamingListener(ExchangeCredentials credentials, Action<object, NotificationEventArgs> onNotificationEvent)
        {
            _onNotificationEvent = onNotificationEvent;
            _exchangeVersion =  ExchangeVersion.Exchange2010_SP1;
            this._credentials = credentials;
        }

        /// <summary>
        /// Creates a new Notification subscription for the desired user and starts listening. Automatically assigns subscriptions to adequate CAS connections. Uses AutoDiscover to determine User's EWS Url.
        /// </summary>
        /// <param name="userMailAddress">The desired user's mail address. Used for AutoDiscover</param>
        /// <param name="folderIds">The Exchange folders under observation</param>
        /// <param name="eventTypes">Notifications will be received for these eventTypes</param>
        public void AddSubscription(MailAddress userMailAddress, IEnumerable<FolderId> folderIds, IEnumerable<EventType> eventTypes)
        {
            var exchangeService = new ExchangeService(this._exchangeVersion) { Credentials = this._credentials };
            exchangeService.AutodiscoverUrl(userMailAddress.ToString(), x => true);
            exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMailAddress.ToString());
            var ewsUrl = exchangeService.Url;

            var collection = FindOrCreateSubscriptionCollection(exchangeService);
            collection.Add(exchangeService.SubscribeToStreamingNotifications(folderIds, eventTypes.ToArray()));
            this._subscriptionCollections.Add(collection);
        }

        private StreamingSubscriptionCollection FindOrCreateSubscriptionCollection(ExchangeService service)
        {
            var collection = _subscriptionCollections.FirstOrDefault(s => s.TargetEwsUrl.ToString() == service.Url.ToString()) ??
                            new StreamingSubscriptionCollection(
                                service,
                                this._onNotificationEvent);
            return collection;
        }
    }
}