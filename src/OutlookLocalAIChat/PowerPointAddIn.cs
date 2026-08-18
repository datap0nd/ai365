using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.UI;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat
{
    // AI365 for PowerPoint: the same restrained chat pane, hosted by
    // PowerPoint. The add-in itself holds no capability beyond
    // opening the sidebar; every read and draft boundary lives in
    // the pane and its tool hosts.
    [ComVisible(true)]
    [Guid("69FAE812-274F-43F8-8F45-1B4EB22B5248")]
    [ProgId("AI365.PowerPointAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PowerPointAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _powerPointApplication;
        private object _ctpFactory;
        private object _taskPane;
        private OfficeChatPane _chatPane;

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _powerPointApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _powerPointApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            CloseTaskPane();
        }

        public void CTPFactoryAvailable(object ctpFactory)
        {
            _ctpFactory = ctpFactory;
        }

        public string GetCustomUI(string ribbonId)
        {
            if (!string.Equals(
                ribbonId,
                "Microsoft.Powerpoint.Presentation",
                StringComparison.Ordinal))
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabHome\">" +
                "<group id=\"AI365.PowerPoint.Group\" label=\"AI365\">" +
                "<button id=\"AI365.PowerPoint.Open\" label=\"AI365\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open AI365\" " +
                "supertip=\"Chat with your presentation. AI365 never saves, deletes, or sends anything.\"/>" +
                "</group></tab></tabs></ribbon>" +
                "</customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_powerPointApplication == null)
                {
                    MessageBox.Show(
                        "PowerPoint is not ready. Restart PowerPoint and try again.",
                        "AI365",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "PowerPoint has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "AI365",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_taskPane == null)
                {
                    dynamic factory = _ctpFactory;
                    _taskPane = factory.CreateCTP(
                        "AI365.OfficePane",
                        "AI365",
                        Type.Missing);

                    dynamic pane = _taskPane;
                    pane.DockPosition = 2;
                    pane.Width = 380;

                    _chatPane = pane.ContentControl as
                        OfficeChatPane ??
                        OfficeChatPane.LastCreated;
                    if (_chatPane == null)
                    {
                        throw new InvalidOperationException(
                            "PowerPoint created the sidebar but its chat control was unavailable.");
                    }

                    _chatPane.Initialize(
                        "powerpoint",
                        _powerPointApplication);
                    pane.Visible = true;
                }
                else
                {
                    dynamic pane = _taskPane;
                    pane.Visible = true;
                }
            }
            catch (Exception exception)
            {
                Log.Error("PowerPointOpenChat", exception);
                MessageBox.Show(
                    DiagnosticDetails.ForException(
                        exception,
                        "SIDEBAR_OPEN_FAILED"),
                    "AI365",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CloseTaskPane()
        {
            try
            {
                _chatPane?.Shutdown();
                if (_taskPane != null)
                {
                    dynamic pane = _taskPane;
                    pane.Visible = false;
                }
            }
            catch (Exception exception)
            {
                Log.Error("PowerPointCloseTaskPane", exception);
            }

            Release(_taskPane);
            Release(_ctpFactory);
            _taskPane = null;
            _ctpFactory = null;
            _chatPane = null;
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
