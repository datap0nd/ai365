using System;
using System.Runtime.InteropServices;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class MessageReader
    {
        private readonly object _outlookApplication;

        public MessageReader(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public MessageSnapshot CaptureCurrent()
        {
            object item = null;
            object inspector = null;
            object explorer = null;
            object selection = null;
            object parent = null;

            try
            {
                dynamic application = _outlookApplication;
                inspector = application.ActiveInspector();
                if (inspector != null)
                {
                    dynamic activeInspector = inspector;
                    item = activeInspector.CurrentItem;
                }

                if (item == null)
                {
                    explorer = application.ActiveExplorer();
                    if (explorer != null)
                    {
                        dynamic activeExplorer = explorer;
                        selection = activeExplorer.Selection;
                        dynamic currentSelection = selection;
                        if (currentSelection != null && currentSelection.Count > 0)
                        {
                            item = currentSelection.Item(1);
                        }
                    }
                }

                if (item == null)
                {
                    throw new InvalidOperationException(
                        "Select or open an email in Outlook first.");
                }

                dynamic mail = item;
                var messageClass = SafeString(() => mail.MessageClass);
                if (!messageClass.StartsWith(
                    "IPM.Note",
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "The selected Outlook item is not an email.");
                }

                parent = SafeObject(() => mail.Parent);
                var storeId = string.Empty;
                if (parent != null)
                {
                    dynamic folder = parent;
                    storeId = SafeString(() => folder.StoreID);
                }

                return new MessageSnapshot(
                    SafeString(() => mail.EntryID),
                    storeId,
                    TextBoundary.PlainText(
                        SafeString(() => mail.Subject),
                        1000),
                    TextBoundary.PlainText(
                        BuildSender(mail),
                        1000),
                    TextBoundary.PlainText(
                        SafeString(() => mail.To),
                        2000),
                    SafeDateTime(() => mail.ReceivedTime),
                    TextBoundary.PlainText(
                        SafeString(() => mail.Body),
                        TextBoundary.MaxMessageBodyCharacters));
            }
            finally
            {
                Release(parent);
                Release(selection);
                Release(explorer);
                Release(inspector);
                Release(item);
            }
        }

        private static string BuildSender(dynamic mail)
        {
            var name = SafeString(() => mail.SenderName);
            var address = SafeString(() => mail.SenderEmailAddress);
            if (name.Length == 0)
            {
                return address;
            }

            if (address.Length == 0 ||
                name.Equals(address, StringComparison.OrdinalIgnoreCase))
            {
                return name;
            }

            return name + " <" + address + ">";
        }

        private static object SafeObject(Func<object> reader)
        {
            try
            {
                return reader();
            }
            catch
            {
                return null;
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime? SafeDateTime(Func<object> reader)
        {
            try
            {
                var value = reader();
                if (value is DateTime)
                {
                    return (DateTime)value;
                }
            }
            catch
            {
            }

            return null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
