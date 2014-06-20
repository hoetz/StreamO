using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;

namespace StreamO
{
    public class GroupIdentifier
    {
        private string _Value;

        public string Value
        {
            get { return this._Value; }
        }

        public GroupIdentifier(string GroupingInformation, Uri externalEwsUri)
        {
            this._Value = string.Format("{0};{1}", GroupingInformation, externalEwsUri.ToString());
        }
    }

    public class StreamingListener : IDisposable
    {
        private readonly ExchangeCredentials _credentials;
        private readonly ExchangeVersion _exchangeVersion;
        private readonly IList<StreamingSubscriptionCollection> _subscriptionCollections = new List<StreamingSubscriptionCollection>();
        private readonly Action<SubscriptionNotificationEventCollection> _onNotificationEvent;

        /// <summary>
        /// Manages Streaming Notifications for Exchange Users. Automatically assigns subscriptions to adequate CAS connections.
        /// </summary>
        /// <param name="credentials">Credentials with permission to impersonate the user mailboxes for all the subscriptions this instance will manage.</param>
        /// <param name="exchangeVersion">The version of the target Exchange server. Must be 2010 SP1 or higher</param>
        /// <param name="onNotificationEvent">The Action to invoke when Notifications arrive</param>
        public StreamingListener(ExchangeCredentials credentials, ExchangeVersion exchangeVersion, Action<SubscriptionNotificationEventCollection> onNotificationEvent)
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
        /// <param name="onNotificationEvent">The Action to invoke when Notifications for any subscription arrive.</param>
        public StreamingListener(ExchangeCredentials credentials, Action<SubscriptionNotificationEventCollection> onNotificationEvent)
        {
            _onNotificationEvent = onNotificationEvent;
            _exchangeVersion = ExchangeVersion.Exchange2010_SP1;
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
            AutodiscoverService autodiscoverService = new AutodiscoverService(this._exchangeVersion);
            autodiscoverService.Credentials = this._credentials;
            autodiscoverService.RedirectionUrlValidationCallback = x => true;
            //only on o365!
            autodiscoverService.EnableScpLookup = false;

            var exchangeService = new ExchangeService(this._exchangeVersion) { Credentials = this._credentials };

            Debug.WriteLine("Autodiscover EWS Url for Subscription User...");
            //exchangeService.AutodiscoverUrl(userMailAddress.ToString(), x => true);

            var response = autodiscoverService.GetUserSettings(userMailAddress.ToString(), UserSettingName.GroupingInformation, UserSettingName.ExternalEwsUrl);
            string extUrl = "";
            string groupInfo = "";
            response.TryGetSettingValue<string>(UserSettingName.ExternalEwsUrl, out extUrl);
            response.TryGetSettingValue<string>(UserSettingName.GroupingInformation, out groupInfo);

            var ewsUrl = new Uri(extUrl);
            exchangeService.Url = ewsUrl;
            var collection = FindOrCreateSubscriptionCollection(exchangeService, new GroupIdentifier(groupInfo, ewsUrl));
            collection.Add(userMailAddress.ToString(), folderIds, eventTypes.ToArray());
            if (_subscriptionCollections.Contains(collection) == false)
                this._subscriptionCollections.Add(collection);
        }

        /// <summary>
        /// Creates a new Notification subscription for the desired user and starts listening on NewMail events for the inbox. Automatically assigns subscriptions to adequate CAS connections. Uses AutoDiscover to determine User's EWS Url.
        /// </summary>
        /// <param name="userMailAddress"></param>
        public void AddSubscription(string userMailAddress)
        {
            this.AddSubscription(userMailAddress, new FolderId[] { WellKnownFolderName.Inbox }, new EventType[] { EventType.NewMail });
        }

        /// <summary>
        /// Cancels the notification subscription for this user.
        /// </summary>
        /// <param name="userMailAddress">The MailAddress of the user to remove</param>
        /// <returns></returns>
        public bool RemoveSubscription(string userMailAddress)
        {
            var collection = FindBy(userMailAddress);
            if (collection != null)
            {
                Debug.WriteLine(string.Format("Closing subscription for {0}", userMailAddress));
                bool success = collection.Remove(userMailAddress);
                if (collection.ActiveUsers.Any() == false)
                {
                    Debug.WriteLine(string.Format("Removing collection for {0}", collection.TargetEwsUrl.ToString()));
                    success = this._subscriptionCollections.Remove(collection);
                }
                return success;
            }
            else
                return false;
        }

        /// <summary>
        /// Creates a new Notification subscription for the desired user and starts listening. Automatically assigns subscriptions to adequate CAS connections. Uses AutoDiscover to determine User's EWS Url.
        /// </summary>
        /// <param name="userMailAddress">The desired user's mail address. Used for AutoDiscover</param>
        /// <param name="folderIds">The Exchange folders under observation</param>
        /// <param name="eventTypes">Notifications will be received for these eventTypes</param>
        public void AddSubscription(string userMailAddress, IEnumerable<FolderId> folderIds, IEnumerable<EventType> eventTypes)
        {
            var mailAddress = new MailAddress(userMailAddress);
            this.AddSubscription(mailAddress, folderIds, eventTypes);
        }

        private StreamingSubscriptionCollection FindOrCreateSubscriptionCollection(ExchangeService service, GroupIdentifier groupIdentifier)
        {
            var collection = _subscriptionCollections.FirstOrDefault(s => s.groupIdentifier.Value == groupIdentifier.Value) ??
                            new StreamingSubscriptionCollection(
                                service,
                                this._onNotificationEvent, groupIdentifier);
            return collection;
        }

        private StreamingSubscriptionCollection FindBy(string userMailAddress)
        {
            var collection = _subscriptionCollections.FirstOrDefault(s => s.ActiveUsers.Contains(new MailAddress(userMailAddress)));
            return collection;
        }

        public void Dispose()
        {
            foreach (var item in _subscriptionCollections)
            {
                item.Dispose();
            }
        }
    }
}