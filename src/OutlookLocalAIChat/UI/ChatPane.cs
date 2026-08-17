/*
THESIS: A restrained Outlook sidebar makes mailbox retrieval and one linked,
human-reviewed draft visible without granting send capability.
OWN-WORLD: Dark charcoal surfaces with a single blue accent, slim inline
activity lines, the model picker anchored at the bottom, and plain text
throughout. High contrast mode defers to system colors.
STORY: Ask the mailbox, watch slim context entries record what loaded, then
deliberately open an unsent draft.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.UI
{
    [ComVisible(true)]
    [Guid("14D24FA1-4342-442F-B68B-B68D7372794C")]
    [ProgId("OutlookLocalAIChat.ChatPane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ChatPane : UserControl
    {
        private const int WorkingSetExpandedHeight = 322;
        private const int WorkingSetCollapsedHeight = 36;

        // Dark, Ollama-client-inspired palette. High contrast mode keeps
        // system colors so accessibility themes always win.
        private static Color OutlookBlue
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(92, 143, 255);
            }
        }

        private static Color TextPrimary
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.WindowText
                    : Color.FromArgb(232, 232, 236);
            }
        }

        private static Color TextSecondary
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(152, 152, 160);
            }
        }

        private static Color ErrorText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.HotTrack
                    : Color.FromArgb(255, 118, 118);
            }
        }

        private static Color Surface
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Window
                    : Color.FromArgb(26, 27, 30);
            }
        }

        private static Color SurfaceMuted
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Control
                    : Color.FromArgb(33, 34, 38);
            }
        }

        private static Color CardSurface
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Window
                    : Color.FromArgb(43, 44, 49);
            }
        }

        private readonly SettingsStore _settingsStore =
            new SettingsStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly List<ChatTurn> _history =
            new List<ChatTurn>();
        private readonly List<MessageSnapshot> _workingMessages =
            new List<MessageSnapshot>();
        private readonly List<ExternalContextDocument> _externalContext =
            new List<ExternalContextDocument>();
        private readonly RichTextBox _transcript =
            new RichTextBox();
        private readonly TextBox _composer = new TextBox();
        private readonly Label _scopeMeta = new Label();
        private readonly ComboBox _modelPicker = new ComboBox();
        private readonly Label _draftState = new Label();
        private readonly Label _status = new Label();
        private readonly Button _send = new Button();
        private readonly Panel _workingSetLayer = new Panel();
        private readonly Label _workingSetHeading = new Label();
        private readonly FlowLayoutPanel _workingSetCards =
            new FlowLayoutPanel();
        private TableLayoutPanel _rootLayout;
        private Button _workingSetToggle;
        private Button _workingSetClear;
        private Button _refresh;
        private Button _addFiles;
        private Button _newChat;
        private Button _settingsButton;

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private DraftToolHost _draftTools;
        private CancellationTokenSource _requestCancellation;
        private bool _busy;
        private bool _shutdown;
        private bool _workingSetExpanded = true;

        public ChatPane()
        {
            LastCreated = this;
            _settings = _settingsStore.Load();

            Dock = DockStyle.Fill;
            BackColor = Surface;
            ForeColor = TextPrimary;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            MinimumSize = new Size(300, 480);
            AllowDrop = true;
            DragEnter += ChatPaneDragEnter;
            DragDrop += ChatPaneDragDrop;
            BuildLayout();
            RefreshModelPicker();
            ShowWelcome();
        }

        internal static ChatPane LastCreated { get; private set; }

        internal void Initialize(
            object outlookApplication,
            bool refreshSelection = true)
        {
            if (_outlookApplication != null)
            {
                return;
            }

            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
            _draftTools = new DraftToolHost(
                _outlookApplication);
            if (refreshSelection)
            {
                RefreshSelectedMessage();
            }
            UpdateDraftState();
            _composer.Focus();
        }

        public void RefreshSelectedMessage()
        {
            if (_outlookApplication == null)
            {
                SetScopeUnavailable(
                    "Outlook is still initializing.");
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Still working - try again in a moment.",
                    true);
                return;
            }

            try
            {
                SetSelectedMessage(
                    new MessageReader(_outlookApplication)
                        .CaptureCurrent());
                SetStatus(
                    "Email selected",
                    false);
            }
            catch (Exception exception)
            {
                _workingMessages.Clear();
                RefreshContextLayer("External files");
                _selectedMessage = null;
                SetScopeUnavailable(
                    "No selected email. Mailbox search is still available.");
                SetStatus(
                    "Ready",
                    false);
                Log.Error("CaptureCurrent", exception);
            }
        }

        public void UseRibbonSelection(object selection)
        {
            if (_outlookApplication == null)
            {
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Still working - try again in a moment.",
                    true);
                return;
            }

            try
            {
                var reader = new MessageReader(
                    _outlookApplication);
                IReadOnlyList<MessageSnapshot> messages;
                try
                {
                    messages = selection == null
                        ? reader.CaptureActiveSelectionMany()
                        : reader.CaptureSelectionMany(selection);
                }
                catch when (selection != null)
                {
                    messages = reader.CaptureActiveSelectionMany();
                }

                ApplySelectedMessages(messages);
            }
            catch (Exception exception)
            {
                Log.Error("CaptureRibbonSelection", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EMAIL_SELECTION_FAILED");
                SetStatus(FirstLine(details), true);
            }
        }

        public void AddActiveSelection()
        {
            UseRibbonSelection(null);
        }

        private void ApplySelectedMessages(
            IReadOnlyList<MessageSnapshot> messages)
        {
            if (messages == null || messages.Count == 0)
            {
                throw new InvalidOperationException(
                    "Select one to ten emails in Outlook first.");
            }

            if (messages.Count == 1)
            {
                SetSelectedMessage(messages[0]);
                SetStatus(
                    "Email added to context",
                    false);
                return;
            }

            SetWorkingMessages(
                messages,
                messages.Count + " emails selected in Outlook");
        }

        private void HandleLocalSearchCommand(
            string prompt,
            LocalSearchCommand command)
        {
            AppendTurn("You", prompt, OutlookBlue);
            _composer.Clear();
            switch (command.Kind)
            {
                case LocalSearchCommandKind.Help:
                    AppendContext(
                        "Use /search <person or topic> to replace the working set with " +
                        "the newest ten matching emails from Inbox and Sent Items. " +
                        "Use /search clear to remove the working set.");
                    SetStatus(
                        "Search help shown",
                        false);
                    return;
                case LocalSearchCommandKind.Clear:
                    _workingMessages.Clear();
                    RefreshContextLayer("External files");
                    _selectedMessage = null;
                    SetScopeUnavailable(
                        "No working set. Use /search or select email in Outlook.");
                    AppendContext(
                        "Working set cleared. No email bodies are loaded.");
                    SetStatus(
                        "Working set cleared",
                        false);
                    return;
                case LocalSearchCommandKind.Search:
                    SearchWorkingMessages(command.Query);
                    return;
                default:
                    return;
            }
        }

        private void SearchWorkingMessages(string query)
        {
            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing.",
                    true);
                return;
            }

            UseWaitCursor = true;
            SetStatus(
                "Searching mailbox...",
                false);
            try
            {
                var hits = new MailboxContextService(
                    _outlookApplication)
                    .Search(
                        query,
                        "all",
                        3650,
                        MailboxWorkingSet.MaxMessages);
                var messages = new List<MessageSnapshot>();
                foreach (var hit in hits)
                {
                    messages.Add(hit.Message);
                }

                if (messages.Count == 0)
                {
                    AppendContext(
                        "No emails matched '" + query + "'. " +
                        (_workingMessages.Count > 0
                            ? "The previous working set was kept."
                            : "No working set was created.") +
                        " Refine the person or topic and run /search again.");
                    SetStatus(
                        "No matches - refine /search",
                        true);
                    return;
                }

                SetWorkingMessages(
                    messages,
                    "Search: " + query);
            }
            catch (Exception exception)
            {
                Log.Error("LocalMailboxSearch", exception);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "LOCAL_SEARCH_FAILED");
                AppendError(details);
                SetStatus(FirstLine(details), true);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void SetWorkingMessages(
            IEnumerable<MessageSnapshot> messages,
            string source)
        {
            var bounded = MailboxWorkingSet.Normalize(messages);
            _workingMessages.Clear();
            foreach (var message in bounded)
            {
                _workingMessages.Add(message);
            }

            _selectedMessage = null;
            _scopeMeta.Text =
                "Working set: " +
                _workingMessages.Count +
                " of " +
                MailboxWorkingSet.MaxMessages +
                " emails";
            ShowWorkingSetLayer(source, _workingMessages);
            AppendContext(
                TextBoundary.SingleLine(source, 260) +
                ". The ten-email context layer is ready. " +
                "Search again to replace it if needed.");
            SetStatus(
                "Working set ready",
                false);
        }

        private void ShowWorkingSetLayer(
            string source,
            IReadOnlyList<MessageSnapshot> messages)
        {
            RefreshContextLayer(source);
        }

        private void RefreshContextLayer(string source)
        {
            ClearWorkingSetCards();
            var emailCount = _workingMessages.Count;
            var fileCount = _externalContext.Count;
            if (emailCount + fileCount == 0)
            {
                HideWorkingSetLayer();
                return;
            }

            _workingSetHeading.Text =
                "Context - " +
                (emailCount + fileCount) +
                ((emailCount + fileCount) == 1
                    ? " item"
                    : " items");
            _workingSetHeading.AccessibleDescription =
                TextBoundary.SingleLine(source, 260) +
                ". Bounded user-approved email and file context.";
            for (var index = 0; index < emailCount; index++)
            {
                _workingSetCards.Controls.Add(
                    BuildWorkingSetCard(
                        index,
                        _workingMessages[index]));
            }

            for (var index = 0; index < fileCount; index++)
            {
                _workingSetCards.Controls.Add(
                    BuildExternalContextCard(
                        index,
                        _externalContext[index]));
            }

            _workingSetExpanded = true;
            _workingSetCards.Visible = true;
            _workingSetToggle.Text = "Hide";
            _workingSetToggle.AccessibleName =
                "Hide working set";
            _workingSetLayer.Visible = true;
            SetWorkingSetRowHeight(
                WorkingSetExpandedHeight);
            ResizeWorkingSetCards();
        }

        private Control BuildExternalContextCard(
            int index,
            ExternalContextDocument document)
        {
            var card = new Panel
            {
                Height = 50,
                Margin = new Padding(0, 0, 0, 6),
                Padding = new Padding(10, 5, 10, 5),
                BackColor = CardSurface,
                AccessibleName = "Context file: " + document.Name,
                AccessibleDescription =
                    document.Content.Length + " bounded text characters"
            };
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = CardSurface,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 36));
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 52));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 48));
            var badge = new Label
            {
                Dock = DockStyle.Fill,
                Text = "FILE",
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = OutlookBlue,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(7F, Font.Size - 2F),
                    FontStyle.Bold)
            };
            var name = new Label
            {
                Dock = DockStyle.Fill,
                Text = document.Name,
                AutoEllipsis = true,
                ForeColor = TextPrimary,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var metadata = new Label
            {
                Dock = DockStyle.Fill,
                Text = document.Content.Length + " text characters",
                AutoEllipsis = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular)
            };
            grid.Controls.Add(badge, 0, 0);
            grid.SetRowSpan(badge, 2);
            grid.Controls.Add(name, 1, 0);
            grid.Controls.Add(metadata, 1, 1);
            card.Controls.Add(grid);
            return card;
        }

        private Control BuildWorkingSetCard(
            int index,
            MessageSnapshot message)
        {
            var subject = TextBoundary.SingleLine(
                SubjectDisplay.Clean(message.Subject),
                180);
            if (subject.Length == 0)
            {
                subject = "(No subject)";
            }

            var sender = TextBoundary.SingleLine(
                message.Sender,
                120);
            if (sender.Length == 0)
            {
                sender = "Unknown sender";
            }

            var date = message.ReceivedAt?.ToString(
                "yyyy-MM-dd HH:mm") ??
                "Unknown date";
            var card = new Panel
            {
                Height = 50,
                Margin = new Padding(0, 0, 0, 6),
                Padding = new Padding(10, 5, 10, 5),
                BackColor = CardSurface,
                AccessibleName =
                    "Email " +
                    (index + 1) +
                    ": " +
                    subject,
                AccessibleDescription =
                    sender + ", " + date
            };
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                BackColor = CardSurface,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 28));
            grid.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 52));
            grid.RowStyles.Add(
                new RowStyle(SizeType.Percent, 48));

            var number = new Label
            {
                Dock = DockStyle.Fill,
                Text = (index + 1).ToString(),
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = OutlookBlue,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var subjectLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = subject,
                AutoEllipsis = true,
                ForeColor = TextPrimary,
                Font = new Font(
                    Font.FontFamily,
                    Font.Size,
                    FontStyle.Bold)
            };
            var metadata = new Label
            {
                Dock = DockStyle.Fill,
                Text = sender + " | " + date,
                AutoEllipsis = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular)
            };
            grid.Controls.Add(number, 0, 0);
            grid.SetRowSpan(number, 2);
            grid.Controls.Add(subjectLabel, 1, 0);
            grid.Controls.Add(metadata, 1, 1);
            card.Controls.Add(grid);
            return card;
        }

        private void WorkingSetToggleClick(
            object sender,
            EventArgs eventArgs)
        {
            SetWorkingSetExpanded(
                !_workingSetExpanded);
        }

        private void SetWorkingSetExpanded(bool expanded)
        {
            _workingSetExpanded = expanded;
            _workingSetCards.Visible = expanded;
            _workingSetToggle.Text =
                expanded ? "Hide" : "Show";
            _workingSetToggle.AccessibleName =
                (expanded ? "Hide" : "Show") +
                " context";
            SetWorkingSetRowHeight(
                expanded
                    ? WorkingSetExpandedHeight
                    : WorkingSetCollapsedHeight);
        }

        private void ResizeWorkingSetCards()
        {
            var width = Math.Max(
                120,
                _workingSetCards.ClientSize.Width -
                SystemInformation.VerticalScrollBarWidth -
                2);
            foreach (Control card in
                _workingSetCards.Controls)
            {
                card.Width = width;
            }
        }

        private void HideWorkingSetLayer()
        {
            ClearWorkingSetCards();
            _workingSetLayer.Visible = false;
            SetWorkingSetRowHeight(0);
        }

        private void ClearContextClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            _workingMessages.Clear();
            _externalContext.Clear();
            _selectedMessage = null;
            HideWorkingSetLayer();
            SetScopeUnavailable(
                "No context selected. Use /search, Add email, or Add files.");
            SetStatus(
                "Context cleared",
                false);
        }

        private void AddFilesClick(
            object sender,
            EventArgs eventArgs)
        {
            using (var dialog = new OpenFileDialog
            {
                Title = "Add bounded text context to MetoMail",
                Multiselect = true,
                CheckFileExists = true,
                Filter =
                    "Supported text files|*.txt;*.md;*.csv;*.json;*.xml;*.html;*.htm;*.log;*.yaml;*.yml;*.ini|" +
                    "All files|*.*"
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    AddExternalFiles(dialog.FileNames);
                }
            }
        }

        private void AddExternalFiles(IEnumerable<string> paths)
        {
            try
            {
                var loaded = ExternalContextLoader.LoadFiles(paths);
                var combined = new List<ExternalContextDocument>(
                    _externalContext);
                combined.AddRange(loaded);
                var normalized =
                    ExternalContextDocument.Normalize(combined);
                _externalContext.Clear();
                foreach (var document in normalized)
                {
                    _externalContext.Add(document);
                }

                RefreshContextLayer("External files");
                SetStatus(
                    _externalContext.Count +
                    (_externalContext.Count == 1
                        ? " external context file is ready."
                        : " external context files are ready.") +
                    " File text is bounded and treated as untrusted data.",
                    false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "EXTERNAL_CONTEXT_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("AddExternalContext", exception);
            }
        }

        private void ChatPaneDragEnter(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                eventArgs.Effect = DragDropEffects.None;
                return;
            }

            if (HasOutlookDragFormat(eventArgs.Data))
            {
                eventArgs.Effect = DragDropEffects.Link;
                return;
            }

            if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
            {
                eventArgs.Effect = DragDropEffects.Copy;
                return;
            }

            eventArgs.Effect = DragDropEffects.None;
        }

        private void ChatPaneDragDrop(
            object sender,
            DragEventArgs eventArgs)
        {
            if (_busy || eventArgs.Data == null)
            {
                return;
            }

            try
            {
                if (HasOutlookDragFormat(eventArgs.Data))
                {
                    AddActiveSelection();
                    return;
                }

                if (eventArgs.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var paths = eventArgs.Data.GetData(
                        DataFormats.FileDrop) as string[];
                    AddExternalFiles(paths);
                    return;
                }

            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "CONTEXT_DROP_FAILED");
                SetStatus(FirstLine(details), true);
                Log.Error("DropContext", exception);
            }
        }

        private static bool HasOutlookDragFormat(IDataObject data)
        {
            foreach (var format in data.GetFormats())
            {
                if (format.Equals(
                        "RenPrivateMessages",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptor",
                        StringComparison.OrdinalIgnoreCase) ||
                    format.Equals(
                        "FileGroupDescriptorW",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void ClearWorkingSetCards()
        {
            while (_workingSetCards.Controls.Count > 0)
            {
                var card = _workingSetCards.Controls[0];
                _workingSetCards.Controls.RemoveAt(0);
                card.Dispose();
            }
        }

        private void SetWorkingSetRowHeight(float height)
        {
            if (_rootLayout == null ||
                _rootLayout.RowStyles.Count < 3)
            {
                return;
            }

            _rootLayout.RowStyles[2].Height = height;
            _rootLayout.PerformLayout();
        }

        internal void Shutdown()
        {
            if (_shutdown)
            {
                return;
            }

            _shutdown = true;
            _requestCancellation?.Cancel();
            _requestCancellation?.Dispose();
            _requestCancellation = null;
            _client.Dispose();
            _draftTools?.Dispose();
            _draftTools = null;
            _outlookApplication = null;
            if (ReferenceEquals(LastCreated, this))
            {
                LastCreated = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Shutdown();
            }

            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            _rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(0),
                BackColor = Surface
            };
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 56));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 34));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 0));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 104));
            _rootLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 52));

            _rootLayout.Controls.Add(BuildHeader(), 0, 0);
            _rootLayout.Controls.Add(BuildToolbar(), 0, 1);
            _rootLayout.Controls.Add(BuildWorkingSetLayer(), 0, 2);
            _rootLayout.Controls.Add(BuildTranscript(), 0, 3);
            _rootLayout.Controls.Add(BuildComposer(), 0, 4);
            _rootLayout.Controls.Add(BuildBottomBar(), 0, 5);
            Controls.Add(_rootLayout);
        }

        private Control BuildWorkingSetLayer()
        {
            _workingSetLayer.Dock = DockStyle.Fill;
            _workingSetLayer.BackColor = SurfaceMuted;
            _workingSetLayer.Padding = new Padding(14, 0, 14, 8);
            _workingSetLayer.Visible = false;
            _workingSetLayer.AccessibleName =
                "Selected email working set";

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = WorkingSetCollapsedHeight,
                BackColor = SurfaceMuted
            };
            _workingSetHeading.Dock = DockStyle.Fill;
            _workingSetHeading.TextAlign =
                ContentAlignment.MiddleLeft;
            _workingSetHeading.Font = new Font(
                Font.FontFamily,
                Font.Size,
                FontStyle.Bold);
            _workingSetHeading.ForeColor = TextPrimary;

            _workingSetToggle = MakeLinkButton("Hide", 52);
            _workingSetToggle.Dock = DockStyle.Right;
            _workingSetToggle.Click += WorkingSetToggleClick;
            _workingSetClear = MakeLinkButton("Clear", 52);
            _workingSetClear.Dock = DockStyle.Right;
            _workingSetClear.Click += ClearContextClick;
            header.Controls.Add(_workingSetHeading);
            header.Controls.Add(_workingSetClear);
            header.Controls.Add(_workingSetToggle);

            _workingSetCards.Dock = DockStyle.Fill;
            _workingSetCards.AutoScroll = true;
            _workingSetCards.FlowDirection =
                FlowDirection.TopDown;
            _workingSetCards.WrapContents = false;
            _workingSetCards.Padding = new Padding(0, 0, 0, 4);
            _workingSetCards.BackColor = SurfaceMuted;
            _workingSetCards.Resize +=
                (sender, args) => ResizeWorkingSetCards();

            _workingSetLayer.Controls.Add(_workingSetCards);
            _workingSetLayer.Controls.Add(header);
            return _workingSetLayer;
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted,
                Padding = new Padding(14, 8, 14, 6),
                ColumnCount = 2,
                RowCount = 2
            };
            header.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 44));
            header.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            header.RowStyles.Add(
                new RowStyle(SizeType.Percent, 50));
            header.RowStyles.Add(
                new RowStyle(SizeType.Percent, 50));

            var logo = new Panel
            {
                Size = new Size(36, 36),
                Margin = new Padding(0, 2, 8, 0),
                BackColor = SurfaceMuted,
                AccessibleName = "MetoMail logo",
                AccessibleRole = AccessibleRole.Graphic
            };
            logo.Paint += PaintLogo;

            _scopeMeta.AutoEllipsis = true;
            _scopeMeta.Dock = DockStyle.Fill;
            _scopeMeta.ForeColor = TextPrimary;
            _scopeMeta.Font = new Font(
                Font.FontFamily,
                Font.Size,
                FontStyle.Bold);
            _scopeMeta.Text =
                "No context - use /search or select emails";

            _draftState.AutoSize = false;
            _draftState.AutoEllipsis = true;
            _draftState.Dock = DockStyle.Fill;
            _draftState.Text =
                "Say 'create a draft' to open one. MetoMail cannot send.";
            _draftState.ForeColor = TextSecondary;
            _draftState.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Regular);
            _draftState.AccessibleName = "Draft safety status";

            header.Controls.Add(logo, 0, 0);
            header.SetRowSpan(logo, 2);
            header.Controls.Add(_scopeMeta, 1, 0);
            header.Controls.Add(_draftState, 1, 1);
            return header;
        }

        private static void PaintLogo(
            object sender,
            PaintEventArgs eventArgs)
        {
            var panel = (Control)sender;
            var graphics = eventArgs.Graphics;
            graphics.SmoothingMode =
                System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            var side = Math.Min(panel.Width, panel.Height) - 1;
            var bounds = new Rectangle(0, 0, side, side);
            using (var path = RoundedRectangle(bounds, side / 5))
            using (var brush = new SolidBrush(OutlookBlue))
            {
                graphics.FillPath(brush, path);
            }

            var envelope = new Rectangle(
                bounds.Left + (int)(side * 0.22),
                bounds.Top + (int)(side * 0.30),
                (int)(side * 0.56),
                (int)(side * 0.40));
            var flapY = envelope.Top + (int)(envelope.Height * 0.52);
            var centerX = envelope.Left + envelope.Width / 2;
            using (var pen = new Pen(
                Color.White,
                Math.Max(2F, side / 16F)))
            {
                pen.LineJoin =
                    System.Drawing.Drawing2D.LineJoin.Round;
                graphics.DrawRectangle(
                    pen,
                    envelope.Left,
                    envelope.Top,
                    envelope.Width,
                    envelope.Height);
                graphics.DrawLine(
                    pen,
                    envelope.Left,
                    envelope.Top,
                    centerX,
                    flapY);
                graphics.DrawLine(
                    pen,
                    centerX,
                    flapY,
                    envelope.Right,
                    envelope.Top);
            }
        }

        private static System.Drawing.Drawing2D.GraphicsPath
            RoundedRectangle(Rectangle bounds, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            var diameter = Math.Max(2, radius * 2);
            path.AddArc(
                bounds.Left,
                bounds.Top,
                diameter,
                diameter,
                180,
                90);
            path.AddArc(
                bounds.Right - diameter,
                bounds.Top,
                diameter,
                diameter,
                270,
                90);
            path.AddArc(
                bounds.Right - diameter,
                bounds.Bottom - diameter,
                diameter,
                diameter,
                0,
                90);
            path.AddArc(
                bounds.Left,
                bounds.Bottom - diameter,
                diameter,
                diameter,
                90,
                90);
            path.CloseFigure();
            return path;
        }

        private Control BuildToolbar()
        {
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(10, 3, 8, 1),
                BackColor = SurfaceMuted
            };

            _refresh = MakeLinkButton("Add email", 74);
            _refresh.Click +=
                (sender, args) => AddActiveSelection();
            _addFiles = MakeLinkButton("Add files", 68);
            _addFiles.Click += AddFilesClick;
            _newChat = MakeLinkButton("New", 44);
            _newChat.Click += NewChatClick;
            _settingsButton = MakeLinkButton("Settings", 64);
            _settingsButton.Click += SettingsClick;

            toolbar.Controls.Add(_refresh);
            toolbar.Controls.Add(_addFiles);
            toolbar.Controls.Add(_newChat);
            toolbar.Controls.Add(_settingsButton);
            return toolbar;
        }

        private Control BuildTranscript()
        {
            _transcript.Dock = DockStyle.Fill;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = Surface;
            _transcript.ForeColor = TextPrimary;
            _transcript.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _transcript.ReadOnly = true;
            _transcript.DetectUrls = false;
            _transcript.HideSelection = false;
            _transcript.ScrollBars =
                RichTextBoxScrollBars.Vertical;
            _transcript.AccessibleName =
                "MetoMail conversation";
            _transcript.AccessibleDescription =
                "Plain-text mailbox conversation and context-loading ledger.";

            var frame = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14, 10, 14, 8),
                BackColor = Surface
            };
            frame.Controls.Add(_transcript);
            return frame;
        }

        private Control BuildComposer()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(14, 8, 14, 4),
                BackColor = Surface
            };
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 92));
            panel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));

            _composer.Dock = DockStyle.Fill;
            _composer.Multiline = true;
            _composer.AcceptsReturn = true;
            _composer.ScrollBars = ScrollBars.Vertical;
            _composer.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _composer.BorderStyle = BorderStyle.FixedSingle;
            _composer.BackColor = CardSurface;
            _composer.ForeColor = TextPrimary;
            _composer.MaxLength =
                TextBoundary.MaxUserPromptCharacters;
            _composer.AccessibleName = "Message to AI";
            _composer.AccessibleDescription =
                "Ask about the mailbox or request draft text. " +
                "Control Enter submits the prompt.";
            _composer.KeyDown += ComposerKeyDown;

            ConfigurePrimaryButton(_send, "Send \u2191");
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(8, 0, 0, 0);
            _send.AccessibleName = "Send message";
            _send.Click += SendClick;

            panel.Controls.Add(_composer, 0, 0);
            panel.Controls.Add(_send, 1, 0);
            return panel;
        }

        private Control BuildBottomBar()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(14, 4, 14, 8),
                BackColor = Surface
            };
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, 170));
            panel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100));

            _modelPicker.Dock = DockStyle.Fill;
            _modelPicker.DropDownStyle =
                ComboBoxStyle.DropDownList;
            _modelPicker.FlatStyle = FlatStyle.Flat;
            _modelPicker.BackColor = CardSurface;
            _modelPicker.ForeColor = TextPrimary;
            _modelPicker.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Regular);
            _modelPicker.DrawMode = DrawMode.OwnerDrawFixed;
            _modelPicker.ItemHeight = Math.Max(
                _modelPicker.ItemHeight,
                _modelPicker.Font.Height + 6);
            _modelPicker.DrawItem += ModelPickerDrawItem;
            _modelPicker.AccessibleName = "Active AI model";
            _modelPicker.AccessibleDescription =
                "Switches the saved model. Vision-tagged models can " +
                "read email images.";
            _modelPicker.SelectionChangeCommitted +=
                ModelPickerChanged;

            _status.Dock = DockStyle.Fill;
            _status.AutoEllipsis = true;
            _status.ForeColor = TextSecondary;
            _status.TextAlign = ContentAlignment.MiddleRight;
            _status.Font = new Font(
                Font.FontFamily,
                Math.Max(8F, Font.Size - 1F),
                FontStyle.Regular);
            _status.Margin = new Padding(8, 0, 0, 0);
            _status.AccessibleName = "Chat status";
            _status.AccessibleRole = AccessibleRole.StatusBar;
            _status.Text = "Reads up to 10 emails - can never send";

            panel.Controls.Add(_modelPicker, 0, 0);
            panel.Controls.Add(_status, 1, 0);
            return panel;
        }

        private void ModelPickerDrawItem(
            object sender,
            DrawItemEventArgs eventArgs)
        {
            eventArgs.DrawBackground();
            if (eventArgs.Index < 0)
            {
                eventArgs.DrawFocusRectangle();
                return;
            }

            var modelId = Convert.ToString(
                _modelPicker.Items[eventArgs.Index]) ??
                string.Empty;
            var isVision = ModelCatalog.IsVisionCapable(modelId);
            var tag = isVision ? "Vision" : "Text";
            var selected =
                (eventArgs.State & DrawItemState.Selected) ==
                DrawItemState.Selected;
            var bounds = eventArgs.Bounds;
            var tagWidth = TextRenderer.MeasureText(
                eventArgs.Graphics,
                tag,
                eventArgs.Font).Width + 8;
            TextRenderer.DrawText(
                eventArgs.Graphics,
                modelId,
                eventArgs.Font,
                new Rectangle(
                    bounds.Left + 2,
                    bounds.Top,
                    Math.Max(16, bounds.Width - tagWidth - 8),
                    bounds.Height),
                selected
                    ? SystemColors.HighlightText
                    : TextPrimary,
                TextFormatFlags.Left |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis |
                TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(
                eventArgs.Graphics,
                tag,
                eventArgs.Font,
                new Rectangle(
                    bounds.Right - tagWidth - 4,
                    bounds.Top,
                    tagWidth,
                    bounds.Height),
                selected
                    ? SystemColors.HighlightText
                    : (isVision ? OutlookBlue : TextSecondary),
                TextFormatFlags.Right |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPrefix);
            eventArgs.DrawFocusRectangle();
        }

        private void ModelPickerChanged(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            var model = Convert.ToString(
                _modelPicker.SelectedItem) ?? string.Empty;
            if (model.Length == 0 ||
                string.Equals(
                    model,
                    _settings.Model,
                    StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                _settings.Model = model;
                _settingsStore.Save(_settings);
                SetStatus(
                    "Model: " + model +
                    (ModelCatalog.IsVisionCapable(model)
                        ? " (vision)"
                        : " (text)"),
                    false);
            }
            catch (Exception exception)
            {
                Log.Error("SwitchModel", exception);
                SetStatus("The model change was not saved.", true);
            }
        }

        private void RefreshModelPicker()
        {
            _modelPicker.BeginUpdate();
            _modelPicker.Items.Clear();
            var current = (_settings?.Model ?? string.Empty).Trim();
            var models = new List<string>(
                _settings?.DiscoveredModels ?? new List<string>());
            if (current.Length > 0 &&
                models.FindIndex(model =>
                    string.Equals(
                        model,
                        current,
                        StringComparison.OrdinalIgnoreCase)) < 0)
            {
                models.Insert(0, current);
            }

            foreach (var model in models)
            {
                if (!ModelCatalog.IsDisallowedModel(model))
                {
                    _modelPicker.Items.Add(model);
                }
            }

            _modelPicker.EndUpdate();
            if (current.Length > 0)
            {
                var index = _modelPicker.FindStringExact(current);
                if (index >= 0)
                {
                    _modelPicker.SelectedIndex = index;
                }
            }
        }

        private async void SendClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                _requestCancellation?.Cancel();
                return;
            }

            var prompt = TextBoundary.PlainText(
                _composer.Text,
                TextBoundary.MaxUserPromptCharacters);
            if (prompt.Length == 0)
            {
                SetStatus(
                    "Type a message first",
                    true);
                return;
            }

            var localCommand = LocalSearchCommand.Parse(prompt);
            if (localCommand.Kind != LocalSearchCommandKind.None)
            {
                HandleLocalSearchCommand(prompt, localCommand);
                return;
            }

            if (_outlookApplication == null)
            {
                SetStatus(
                    "[OUTLOOK_NOT_READY] Outlook is still initializing.",
                    true);
                return;
            }

            if (!_settings.IsConfigured)
            {
                OpenSettings();
                if (!_settings.IsConfigured)
                {
                    return;
                }
            }

            var requestSelectedMessage = _selectedMessage;
            var requestWorkingMessages =
                new List<MessageSnapshot>(_workingMessages);
            var requestExternalContext =
                new List<ExternalContextDocument>(_externalContext);
            var hasLinkedDraft =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    !hasLinkedDraft &&
                    DraftIntentPolicy.AllowsCreate(prompt),
                    hasLinkedDraft &&
                    DraftIntentPolicy.AllowsUpdate(prompt));
            var transcriptStart = _transcript.TextLength;
            AppendTurn("You", prompt, OutlookBlue);
            if (_workingMessages.Count > 0 &&
                _workingSetExpanded)
            {
                SetWorkingSetExpanded(false);
            }
            _composer.Clear();
            SetBusy(true);
            _requestCancellation =
                new CancellationTokenSource();

            try
            {
                var response = await CompleteMailboxChatAsync(
                    requestSelectedMessage,
                    requestWorkingMessages,
                    requestExternalContext,
                    prompt,
                    draftAuthorization,
                    _requestCancellation.Token);

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                AppendTurn(
                    "MetoMail",
                    response,
                    TextPrimary);
                if (draftAuthorization.IsCreated)
                {
                    SetStatus(
                        "Draft created - unsent, open for review",
                        false);
                }
                else if (draftAuthorization.IsUpdated)
                {
                    SetStatus(
                        "Draft updated",
                        false);
                }
                else if (draftAuthorization.IsConsumed)
                {
                    SetStatus(
                        draftAuthorization.CanUpdate
                            ? "Draft update did not complete"
                            : "Draft creation did not complete",
                        true);
                }
                else if (draftAuthorization.CanCreate)
                {
                    SetStatus(
                        "No draft was created",
                        false);
                }
                else
                {
                    SetStatus(
                        hasLinkedDraft
                            ? "Done - draft unchanged"
                            : "Done",
                        false);
                }
            }
            catch (OperationCanceledException)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                SetStatus("Stopped - prompt restored", false);
            }
            catch (Exception exception)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                var details = DiagnosticDetails.ForException(
                    exception,
                    "AI_REQUEST_FAILED");
                AppendError(details);
                SetStatus(
                    FirstLine(details),
                    true);
                Log.Error("CompleteMailboxChat", exception);
            }
            finally
            {
                _requestCancellation?.Dispose();
                _requestCancellation = null;
                SetBusy(false);
                UpdateDraftState();
            }
        }

        private async Task<string> CompleteMailboxChatAsync(
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages,
            IReadOnlyList<ExternalContextDocument> externalContext,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            CancellationToken cancellationToken)
        {
            var activeDraft = draftAuthorization.CanUpdate
                ? _draftTools?.ActiveDraft
                : null;
            var imagesExpected = ModelRouting.ContextMayIncludeImages(
                selectedMessage,
                workingMessages);
            var activeModel = ModelRouting.ResolveForRequest(
                _settings,
                imagesExpected);
            if (ModelRouting.IsTemporaryVisionSwitch(
                    _settings,
                    activeModel))
            {
                SetStatus(
                    "Using " + activeModel + " for images",
                    false);
            }
            else if (imagesExpected &&
                     !ModelCatalog.IsVisionCapable(activeModel))
            {
                SetStatus(
                    activeModel + " is text-only - images will not " +
                    "be read",
                    false);
            }

            var request = ChatRequestFactory.Create(
                activeModel,
                selectedMessage,
                _history,
                prompt,
                draftAuthorization.CanCreate,
                activeDraft,
                draftAuthorization.CanUpdate,
                workingMessages,
                externalContext,
                _settings.UseToneProfile
                    ? _settings.ToneProfile
                    : null);
            var mailboxTools = new MailboxToolHost(
                _outlookApplication,
                selectedMessage,
                workingMessages);
            if (VisionImagePrefetch.TryInject(
                    request,
                    activeModel,
                    mailboxTools,
                    selectedMessage,
                    workingMessages))
            {
                SetStatus(
                    "Images attached for vision",
                    false);
            }
            for (var round = 0;
                 round <= TextBoundary.MaxToolRounds;
                 round++)
            {
                var response = await _client.CompleteAsync(
                    _settings,
                    request,
                    cancellationToken);
                var toolCalls = response.tool_calls;
                if (toolCalls == null || toolCalls.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(response.content))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The model stopped without returning text.");
                    }

                    return response.content;
                }

                if (round == TextBoundary.MaxToolRounds)
                {
                    throw new AiEndpointException(
                        "TOOL_ROUND_LIMIT",
                        "The model exceeded the maximum number of bounded tool rounds.");
                }

                if (toolCalls.Count >
                    TextBoundary.MaxToolCallsPerRound)
                {
                    throw new AiEndpointException(
                        "TOOL_CALL_LIMIT",
                        "The model requested too many tools in one round.");
                }

                var results = new List<MailboxToolResult>();
                foreach (var toolCall in toolCalls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var isDraftCall =
                        DraftToolCatalog.IsDraftTool(
                            toolCall?.function?.name);
                    var result = isDraftCall
                        ? _draftTools.Execute(
                            toolCall,
                            mailboxTools.ResolveHandle,
                            draftAuthorization,
                            toolCalls.Count == 1)
                        : mailboxTools.Execute(toolCall);
                    results.Add(result);
                    if (isDraftCall)
                    {
                        AppendDraftAction(
                            result.StatusText);
                    }
                    else
                    {
                        AppendContext(
                            result.StatusText);
                    }
                    SetStatus(result.StatusText, false);
                }

                activeModel = ModelRouting.ResolveForRequest(
                    _settings,
                    imagesExpected,
                    results);
                request.model = TextBoundary.PlainText(activeModel, 200);
                if (ModelRouting.IsTemporaryVisionSwitch(
                        _settings,
                        activeModel))
                {
                    SetStatus(
                        "Temporarily using " + activeModel +
                        " for loaded email images.",
                        false);
                }

                ChatRequestFactory.AppendToolExchange(
                    request,
                    response,
                    results,
                    activeModel);
            }

            throw new AiEndpointException(
                "TOOL_ROUND_LIMIT",
                "The model did not finish after bounded tool use.");
        }

        private void NewChatClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            _history.Clear();
            _workingMessages.Clear();
            _externalContext.Clear();
            HideWorkingSetLayer();
            _draftTools?.Dispose();
            _draftTools = _outlookApplication == null
                ? null
                : new DraftToolHost(_outlookApplication);
            _transcript.Clear();
            RefreshSelectedMessage();
            ShowWelcome();
            UpdateDraftState();
            SetStatus(
                "New chat",
                false);
            _composer.Focus();
        }

        private void SettingsClick(
            object sender,
            EventArgs eventArgs)
        {
            OpenSettings();
        }

        private void OpenSettings()
        {
            using (var settingsWindow =
                new SettingsWindow(
                    _settingsStore,
                    _settings,
                    _outlookApplication))
            {
                if (settingsWindow.ShowDialog(this) ==
                    DialogResult.OK)
                {
                    _settings =
                        settingsWindow.SavedSettings;
                    RefreshModelPicker();
                    SetStatus(
                        "Settings saved - " + _settings.Model,
                        false);
                }
            }
        }

        private void AppendContext(string text)
        {
            AppendActivityLine(text, TextSecondary);
        }

        private void AppendDraftAction(string text)
        {
            AppendActivityLine(text, OutlookBlue);
        }

        // A single dim activity line, like a modern chat client's
        // inline tool/status entries.
        private void AppendActivityLine(string text, Color color)
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                Math.Max(7F, SystemFonts.MessageBoxFont.Size - 1F),
                FontStyle.Regular);
            _transcript.SelectionColor = color;
            _transcript.AppendText(
                "\u2022 " +
                TextBoundary.SingleLine(text, 400) +
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendError(string text)
        {
            AppendStyledBlock(
                "Error",
                text,
                ErrorText,
                FontStyle.Regular);
        }

        private void AppendStyledBlock(
            string label,
            string text,
            Color color,
            FontStyle bodyStyle)
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            _transcript.SelectionColor = color;
            _transcript.AppendText(label + Environment.NewLine);
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                bodyStyle);
            _transcript.SelectionColor = color;
            _transcript.AppendText(
                TextBoundary.PlainText(text, 2400) +
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendTurn(
            string speaker,
            string text,
            Color headingColor)
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            _transcript.SelectionColor = headingColor;
            _transcript.AppendText(
                speaker + Environment.NewLine);
            if (speaker != "You")
            {
                AppendFormattedAssistantText(text);
            }
            else
            {
                SetTranscriptBodyStyle(FontStyle.Regular);
                _transcript.AppendText(
                    TextBoundary.PlainText(
                        text,
                        TextBoundary.MaxUserPromptCharacters));
            }

            SetTranscriptBodyStyle(FontStyle.Regular);
            _transcript.AppendText(
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendFormattedAssistantText(string text)
        {
            var formatted = SafeModelText.Format(
                text,
                TextBoundary.MaxAssistantCharacters);
            var position = 0;
            foreach (var range in formatted.BoldRanges)
            {
                if (range.Start > position)
                {
                    SetTranscriptBodyStyle(
                        FontStyle.Regular);
                    _transcript.AppendText(
                        formatted.PlainText.Substring(
                            position,
                            range.Start - position));
                }

                SetTranscriptBodyStyle(FontStyle.Bold);
                _transcript.AppendText(
                    formatted.PlainText.Substring(
                        range.Start,
                        range.Length));
                position = range.Start + range.Length;
            }

            if (position < formatted.PlainText.Length)
            {
                SetTranscriptBodyStyle(FontStyle.Regular);
                _transcript.AppendText(
                    formatted.PlainText.Substring(position));
            }
        }

        private void SetTranscriptBodyStyle(FontStyle style)
        {
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                style);
            _transcript.SelectionColor = TextPrimary;
        }

        private void ScrollTranscript()
        {
            _transcript.SelectionStart =
                _transcript.TextLength;
            _transcript.ScrollToCaret();
        }

        private void RestoreFailedPrompt(
            string prompt,
            int transcriptStart)
        {
            if (_transcript.TextLength > transcriptStart)
            {
                _transcript.Select(
                    transcriptStart,
                    _transcript.TextLength -
                    transcriptStart);
                _transcript.SelectedText =
                    string.Empty;
            }

            _composer.Text = prompt;
            _composer.SelectionStart =
                _composer.TextLength;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            _send.Text = busy ? "Stop" : "Send \u2191";
            _modelPicker.Enabled = !busy;
            _composer.Enabled = !busy;
            _refresh.Enabled = !busy;
            _addFiles.Enabled = !busy;
            _newChat.Enabled = !busy;
            _settingsButton.Enabled = !busy;
            if (busy)
            {
                SetStatus(
                    "Thinking...",
                    false);
            }
        }

        private void UpdateDraftState()
        {
            var linked =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            _draftState.Text = linked
                ? "Draft linked - feedback updates it. MetoMail cannot send."
                : "Say 'create a draft' to open one. MetoMail cannot send.";
            _draftState.ForeColor = linked
                ? OutlookBlue
                : TextSecondary;
            var style = linked
                ? FontStyle.Bold
                : FontStyle.Regular;
            if (_draftState.Font.Style != style)
            {
                var previousFont = _draftState.Font;
                _draftState.Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    style);
                previousFont.Dispose();
            }
            _draftState.Visible = true;
        }

        private void SetStatus(string text, bool error)
        {
            _status.Text =
                TextBoundary.PlainText(text, 600);
            _status.ForeColor =
                error ? ErrorText : TextSecondary;
        }

        private void SetScopeUnavailable(string text)
        {
            _scopeMeta.Text = text;
        }

        private void SetSelectedMessage(MessageSnapshot message)
        {
            _workingMessages.Clear();
            RefreshContextLayer("External files");
            _selectedMessage = message ??
                throw new ArgumentNullException(nameof(message));
            var displaySubject = SubjectDisplay.Clean(
                _selectedMessage.Subject);
            _scopeMeta.Text =
                "Selected: " +
                (string.IsNullOrWhiteSpace(displaySubject)
                    ? "(No subject)"
                    : displaySubject);
        }

        private void ShowWelcome()
        {
            AppendStyledBlock(
                "MetoMail",
                "Chat with your mailbox - local models, nothing is " +
                "ever sent.\n\n" +
                "\u2022 /search a person or topic\n" +
                "\u2022 Select emails, then Add email (up to 10)\n" +
                "\u2022 Drag emails or files here\n" +
                "\u2022 Ask about attachments, images, invites\n" +
                "\u2022 Say 'create a draft' to write one - it opens " +
                "unsent",
                TextSecondary,
                FontStyle.Regular);
            _transcript.SelectionStart = 0;
            _transcript.ScrollToCaret();
        }

        private void ComposerKeyDown(
            object sender,
            KeyEventArgs eventArgs)
        {
            if (eventArgs.Control &&
                eventArgs.KeyCode == Keys.Enter)
            {
                eventArgs.SuppressKeyPress = true;
                SendClick(_send, EventArgs.Empty);
            }
        }

        private static string FirstLine(string value)
        {
            var text = value ?? string.Empty;
            var index = text.IndexOfAny(
                new[] { '\r', '\n' });
            return index >= 0
                ? text.Substring(0, index)
                : text;
        }

        private static Button MakeLinkButton(
            string text,
            int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                BackColor = SurfaceMuted,
                ForeColor = TextSecondary,
                UseVisualStyleBackColor = false,
                Margin = new Padding(0, 0, 4, 0),
                AccessibleName = text
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = CardSurface;
            button.MouseEnter += (sender, args) =>
                button.ForeColor = TextPrimary;
            button.MouseLeave += (sender, args) =>
                button.ForeColor = TextSecondary;
            return button;
        }

        private static void ConfigurePrimaryButton(
            Button button,
            string text)
        {
            button.Text = text;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = OutlookBlue;
            button.ForeColor =
                SystemColors.HighlightText;
            button.UseVisualStyleBackColor = false;
            button.FlatAppearance.BorderColor =
                OutlookBlue;
            button.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            button.AccessibleName = text;
        }

    }
}
