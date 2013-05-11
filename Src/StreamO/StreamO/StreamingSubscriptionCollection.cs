using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO
{
    internal class StreamingSubscriptionCollection : ICollection<StreamingSubscription>
    {
        private StreamingSubscriptionConnection _connection;
        private static readonly object _conLock = new object();
        private readonly ExchangeService _exchangeService;
        private readonly IList<StreamingSubscription> _subscriptions = new List<StreamingSubscription>();

        /// <summary>
        /// The Url used to call into Exchange Web Services.
        /// </summary>
        public Uri TargetEwsUrl
        {
            get { return _exchangeService.Url; }
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
        /// Adds the <see cref="StreamingSubscription"/> and immediately starts listening. Closes and restarts open connections to EWS. Attention: Use only for subscriptions on the same CAS Server.
        /// </summary>
        /// <param name="item"></param>
        public void Add(StreamingSubscription item)
        {
            lock (_conLock)
            {
                if (_connection.IsOpen)
                    _connection.Close();

                _connection.AddSubscription(item);
                this._subscriptions.Add(item);
                _connection.Open();
            }
        }

        /// <summary>
        /// Removes the <see cref="StreamingSubscription"/> and starts listening again only if any other subscriptions are present.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Remove(StreamingSubscription item)
        {
            bool success;
            lock (_conLock)
            {
                if (_connection.IsOpen)
                    _connection.Close();

                _connection.RemoveSubscription(item);
                success = this._subscriptions.Remove(item);
                if (this._subscriptions.Any())
                    _connection.Open();
            }
            return success;
        }

        public void Clear()
        {
            lock (_conLock)
            {
                _connection.Close();
                _subscriptions.Clear();
            }
        }

        public bool Contains(StreamingSubscription item)
        {
            return this._subscriptions.Contains(item);
        }

        public void CopyTo(StreamingSubscription[] array, int arrayIndex)
        {
            _subscriptions.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get { return _subscriptions.Count; }
        }

        public bool IsReadOnly
        {
            get { return _subscriptions.IsReadOnly; }
        }

        public IEnumerator<StreamingSubscription> GetEnumerator()
        {
            return _subscriptions.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _subscriptions.GetEnumerator();
        }

        private StreamingSubscriptionConnection CreateConnection(Action<object, NotificationEventArgs> OnNotificationEvent)
        {
            var con = new StreamingSubscriptionConnection(this._exchangeService, 30);
            con.OnSubscriptionError += OnSubscriptionError;
            con.OnDisconnect += con_OnDisconnect;

            con.OnNotificationEvent +=
                        new StreamingSubscriptionConnection.NotificationEventDelegate(OnNotificationEvent);

            return con;
        }

        private void con_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            throw new NotImplementedException();
        }

        private void OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            throw new NotImplementedException();
        }
    }
}