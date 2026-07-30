using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.UI;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat
{
    [ComVisible(true)]
    [Guid("0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7")]
    [ProgId("OutlookLocalAIChat.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IAddInCallbacks))]
    public sealed class AddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        IAddInCallbacks
    {
        private object _outlookApplication;
        private ChatWindow _chatWindow;

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _outlookApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseWindow();
            _outlookApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            CloseWindow();
        }

        public string GetCustomUI(string ribbonId)
        {
            var tabId = GetTabId(ribbonId);
            if (tabId == null)
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"" + tabId + "\">" +
                "<group id=\"OutlookLocalAIChat.Group\" label=\"AI Chat\">" +
                "<button id=\"OutlookLocalAIChat.Open\" label=\"Open AI Chat\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Local AI Chat\" " +
                "supertip=\"Discuss the selected email and open unsent drafts.\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_outlookApplication == null)
                {
                    MessageBox.Show(
                        "Outlook is not ready. Restart Outlook and try again.",
                        "Outlook Local AI Chat",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_chatWindow == null || _chatWindow.IsDisposed)
                {
                    _chatWindow = new ChatWindow(_outlookApplication);
                    _chatWindow.FormClosed += (sender, args) =>
                    {
                        _chatWindow = null;
                    };
                    _chatWindow.Show();
                }
                else
                {
                    _chatWindow.RefreshCurrentMessage();
                    if (_chatWindow.WindowState == FormWindowState.Minimized)
                    {
                        _chatWindow.WindowState = FormWindowState.Normal;
                    }

                    _chatWindow.Show();
                    _chatWindow.Activate();
                }
            }
            catch (Exception exception)
            {
                Log.Error("OnOpenChat", exception);
                MessageBox.Show(
                    "AI Chat could not open. Restart Outlook and try again.",
                    "Outlook Local AI Chat",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string GetTabId(string ribbonId)
        {
            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Explorer",
                StringComparison.Ordinal))
            {
                return "TabMail";
            }

            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Mail.Read",
                StringComparison.Ordinal))
            {
                return "TabReadMessage";
            }

            return null;
        }

        private void CloseWindow()
        {
            if (_chatWindow != null && !_chatWindow.IsDisposed)
            {
                _chatWindow.Close();
            }

            _chatWindow = null;
        }
    }
}
