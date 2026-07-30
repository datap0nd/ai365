using System;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class MessageSnapshot
    {
        public MessageSnapshot(
            string entryId,
            string storeId,
            string subject,
            string sender,
            string recipients,
            DateTime? receivedAt,
            string body)
        {
            EntryId = entryId ?? string.Empty;
            StoreId = storeId ?? string.Empty;
            Subject = subject ?? string.Empty;
            Sender = sender ?? string.Empty;
            Recipients = recipients ?? string.Empty;
            ReceivedAt = receivedAt;
            Body = body ?? string.Empty;
        }

        public string EntryId { get; }

        public string StoreId { get; }

        public string Subject { get; }

        public string Sender { get; }

        public string Recipients { get; }

        public DateTime? ReceivedAt { get; }

        public string Body { get; }

        public bool CanReply
        {
            get { return EntryId.Length > 0; }
        }
    }
}
