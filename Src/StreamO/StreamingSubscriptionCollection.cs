using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Mail;

namespace StreamO
{
    internal class StreamingSubscriptionCollection : IDisposable
    {
        private StreamingSubscriptionConnection _connection;
        private static readonly object _conLock = new object();
        private readonly ExchangeService _exchangeService;
        private readonly IDictionary<MailAddress, StreamingSubscription> _subscriptions = new Dictionary<MailAddress, StreamingSubscription>();
        private bool isClosingControlled;

        /// <summary>
        /// The Url used to call into Exchange Web Services.
        /// </summary>
        public Uri TargetEwsUrl
        {
            get { return _exchangeService.Url; }
        }

        public IEnumerable<MailAddress> ActiveUsers
        {
            get { return _subscriptions.Select(x => x.Key); }
        }

        /// <summary>
        /// Manages the connection for multiple <see cref="StreamingSubscription"/> items. Attention: Use only for subscriptions on the same CAS.
        /// </summary>
        /// <param name="exchangeService">The ExchangeService instance this collection uses to connect to the server.</param>
        public StreamingSubscriptionCollection(ExchangeService exchangeService, Action<object, NotificationEventArgs> OnNotificationEvent)
        {
            this._exchangeService = exchangeService;
            _connection = CreateConnection(OnNotificationEvent);
        }

        /// <summary>
        /// Adds the user to subscriptions and starts listening with defined parameters.
        /// </summary>
        /// <param name="userMailAddress">The desired user's mail address.</param>
        /// <param name="folderIds">The Exchange folders under observation</param>
        /// <param name="eventTypes">Notifications will be received for these eventTypes</param>
        public void Add(string userMailAddress, IEnumerable<FolderId> folderIds, IEnumerable<EventType> eventTypes)
        {
            lock (_conLock)
            {
                this.isClosingControlled = true;

                if (_connection.IsOpen)
                    _connection.Close();

                this._exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userMailAddress);

                var item = this._exchangeService.SubscribeToStreamingNotifications(folderIds, eventTypes.ToArray());

                _connection.AddSubscription(item);

                this._subscriptions.Add(new MailAddress(userMailAddress),item);

                _connection.Open();
                Debug.WriteLine(string.Format("Subscription added to EWS connection {0}. Started listening.", this.TargetEwsUrl.ToString()));

                this.isClosingControlled = false;
            }
        }

        /// <summary>
        /// Removes the <see cref="StreamingSubscription"/> and starts listening again only if any other subscriptions are present.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Remove(string userMailAddress)
        {
            bool success;
            var mailAddress=new MailAddress(userMailAddress);
            if (this._subscriptions.ContainsKey(mailAddress) == false)
                return false;

            lock (_conLock)
            {
                this.isClosingControlled = true;
                var subscriptionToRemove = _subscriptions[mailAddress];

                if (_connection.IsOpen)
                    _connection.Close();

                _connection.RemoveSubscription(subscriptionToRemove);
                success = _subscriptions.Remove(mailAddress);
                if (this._subscriptions.Any())
                    _connection.Open();

                this.isClosingControlled = false;
            }
            return success;
        }

        public void Clear()
        {
            lock (_conLock)
            {
                this.isClosingControlled = true;
                _connection.Close();
                _subscriptions.Clear();
                this.isClosingControlled = false;
            }
        }

        private StreamingSubscriptionConnection CreateConnection(Action<object, NotificationEventArgs> OnNotificationEvent)
        {
            var con = new StreamingSubscriptionConnection(this._exchangeService, 30);
            con.OnSubscriptionError += OnSubscriptionError;
            con.OnDisconnect += OnDisconnect;
            
            con.OnNotificationEvent +=
                        new StreamingSubscriptionConnection.NotificationEventDelegate(OnNotificationEvent);

            return con;
        }

        private void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            if (isClosingControlled == false)
            {
                Debug.WriteLine(string.Format("Restoring connection for subscription collection {0}", this.TargetEwsUrl.ToString()));
                this._connection.Open();
            }
        }

        private void OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            lock (_conLock)
            {
                isClosingControlled = true;
                if (_connection.IsOpen)
                    _connection.Close();
                _connection.Dispose();
                isClosingControlled = false;
            }
        }
    }
}