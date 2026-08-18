using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.UI;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat
{
    // AI365 for Excel: the same restrained chat pane, hosted by
    // Excel. The add-in itself holds no capability beyond opening
    // the sidebar; every read and draft boundary lives in the pane
    // and its tool hosts.
    [ComVisible(true)]
    [Guid("C0ABFA36-9854-434D-A542-DD834938737F")]
    [ProgId("AI365.ExcelAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _excelApplication;
        private object _ctpFactory;
        private object _taskPane;
        private OfficeChatPane _chatPane;

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _excelApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _excelApplication = null;
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
            // This add-in is registered only under Excel, so any
            // ribbon id it receives is Excel's; exact-id matching
            // would silently hide the button on a casing change.
            if (string.IsNullOrEmpty(ribbonId))
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabHome\">" +
                "<group id=\"AI365.Excel.Group\" label=\"AI365\">" +
                "<button id=\"AI365.Excel.Open\" label=\"AI365\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open AI365\" " +
                "supertip=\"Chat with your workbook. AI365 never saves, deletes, or sends anything.\"/>" +
                "</group></tab></tabs></ribbon>" +
                "</customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_excelApplication == null)
                {
                    MessageBox.Show(
                        "Excel is not ready. Restart Excel and try again.",
                        "AI365",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "Excel has not made the sidebar service available yet. " +
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
                            "Excel created the sidebar but its chat control was unavailable.");
                    }

                    _chatPane.Initialize(
                        "excel",
                        _excelApplication);
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
                Log.Error("ExcelOpenChat", exception);
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
                Log.Error("ExcelCloseTaskPane", exception);
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
