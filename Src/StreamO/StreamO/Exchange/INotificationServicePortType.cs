using System;
using System.ComponentModel;
using System.ServiceModel;
using System.Xml.Serialization;

namespace StreamO.Exchange
{
    [ServiceContract(Namespace = "http://schemas.microsoft.com/exchange/services/2006/messages",
    ConfigurationName = "Microsoft.Exchange.Notifications.INotificationServicePortType")]
    public interface INotificationServicePortType
    {
        [OperationContract(Action = "*", ReplyAction = "*")]
        [XmlSerializerFormat]
        [ServiceKnownType(typeof(RecurrenceRangeBaseType))]
        [ServiceKnownType(typeof(RecurrencePatternBaseType))]
        [ServiceKnownType(typeof(AttachmentType))]
        [ServiceKnownType(typeof(BasePermissionType))]
        [ServiceKnownType(typeof(BaseItemIdType))]
        [ServiceKnownType(typeof(BaseEmailAddressType))]
        [ServiceKnownType(typeof(BaseFolderIdType))]
        [ServiceKnownType(typeof(BaseFolderType))]
        [ServiceKnownType(typeof(BaseResponseMessageType))]
        SendNotificationResponse SendNotification(SendNotificationRequest request);
    }

    [MessageContract(IsWrapped = false)]
    public class SendNotificationRequest
    {
        [MessageBodyMember(Namespace = "http://schemas.microsoft.com/exchange/services/2006/messages", Order = 0)]
        public SendNotificationResponseType SendNotification;

        public SendNotificationRequest()
        {
        }

        public SendNotificationRequest(SendNotificationResponseType SendNotification)
        {
            this.SendNotification = SendNotification;
        }
    }

    [MessageContract(IsWrapped = false)]
    public class SendNotificationResponse
    {
        [MessageBodyMember(Namespace = "http://schemas.microsoft.com/exchange/services/2006/messages", Order = 0)]
        public
            SendNotificationResultType SendNotificationResult;

        public SendNotificationResponse()
        {
        }

        public SendNotificationResponse(SendNotificationResultType SendNotificationResult)
        {
            this.SendNotificationResult = SendNotificationResult;
        }
    }

    [Serializable]
    [DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.microsoft.com/exchange/services/2006/messages")]
    public class SendNotificationResultType : object, INotifyPropertyChanged
    {
        private SubscriptionStatusType subscriptionStatusField;

        /// <remarks/>
        [XmlElement(Order = 0)]
        public SubscriptionStatusType SubscriptionStatus
        {
            get { return subscriptionStatusField; }
            set
            {
                subscriptionStatusField = value;
                RaisePropertyChanged("SubscriptionStatus");
            }
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        protected void RaisePropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler propertyChanged = PropertyChanged;
            if ((propertyChanged != null))
            {
                propertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/exchange/services/2006/types")]
    public enum SubscriptionStatusType
    {
        /// <remarks/>
        OK,

        /// <remarks/>
        Unsubscribe,
    }
}
