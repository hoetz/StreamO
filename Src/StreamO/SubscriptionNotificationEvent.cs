using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;

namespace StreamO
{
    public class SubscriptionNotificationEventCollection
    {
        private MailAddress _Sender;
        public MailAddress Sender
        {
            get { return this._Sender; }
        }

        private IList<NotificationEvent> _Events = new List<NotificationEvent>();
        public IEnumerable<NotificationEvent> Events
        {
            get { return this._Events; }
        }

        public SubscriptionNotificationEventCollection(MailAddress Sender, IEnumerable<NotificationEvent> Events)
        {
            this._Sender = Sender;
            foreach (var item in Events)
            {
                this._Events.Add(item);
            }
        }
        
    }
}
