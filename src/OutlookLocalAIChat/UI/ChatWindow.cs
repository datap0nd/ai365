/*
THESIS: A restrained Outlook utility makes the read, discuss, and draft boundary
visible, refusing both a generic chat-bubble clone and an autonomous-agent control room.
OWN-WORLD: Windows white and cool-gray surfaces, Outlook blue only for direct actions,
square native fields, one selected-message strip, and plain text throughout.
STORY: Confirm the selected message, converse, then deliberately open an unsent draft.
FIRST VIEWPORT: Message identity at top, transcript in the center, bounded composer and
local draft actions at bottom. The primary action is Send to AI beside the composer.
FORM: Familiar native Outlook utility, pinned by the user's simplicity requirement.
Seed key 61ebccb2. FINISH: unreviewed and undocumented is unfinished; this build ends
with the finish review, the verdict, and DESIGN.md
*/

using System;
using System.Collections.Generic;
using System.Drawing;
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
    public sealed class ChatWindow : Form
    {
        private static Color OutlookBlue
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(0, 95, 184);
            }
        }

        private static Color TextPrimary
        {
            get { return SystemColors.WindowText; }
        }

        private static Color TextSecondary
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(80, 80, 80);
            }
        }

        private static Color Border
        {
            get { return SystemColors.ControlDark; }
        }

        private static Color SurfaceMuted
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Control
                    : Color.FromArgb(244, 246, 248);
            }
        }

        private readonly object _outlookApplication;
        private readonly SettingsStore _settingsStore = new SettingsStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly List<ChatTurn> _history = new List<ChatTurn>();
        private readonly RichTextBox _transcript = new RichTextBox();
        private readonly TextBox _composer = new TextBox();
        private readonly Label _messageTitle = new Label();
        private readonly Label _messageMeta = new Label();
        private readonly Label _status = new Label();
        private readonly Button _send = new Button();
        private readonly Button _replyDraft = new Button();
        private readonly Button _newDraft = new Button();
        private Button _refresh;
        private Button _settingsButton;

        private AppSettings _settings;
        private MessageSnapshot _message;
        private string _lastAssistantText = string.Empty;
        private CancellationTokenSource _requestCancellation;
        private bool _busy;

        public ChatWindow(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
            _settings = _settingsStore.Load();

            Text = "Outlook Local AI Chat";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(540, 720);
            MinimumSize = new Size(440, 580);
            BackColor = SystemColors.Window;
            ForeColor = TextPrimary;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            ShowIcon = false;
            KeyPreview = true;
            KeyDown += WindowKeyDown;
            FormClosed += WindowClosed;

            BuildLayout();
            RefreshCurrentMessage();
        }

        public void RefreshCurrentMessage()
        {
            if (_busy)
            {
                _requestCancellation?.Cancel();
                SetStatus(
                    "Cancelling the current request. Refresh the email again when it stops.",
                    false);
                return;
            }

            try
            {
                var nextMessage =
                    new MessageReader(_outlookApplication).CaptureCurrent();
                var changed =
                    _message == null ||
                    !_message.EntryId.Equals(
                        nextMessage.EntryId,
                        StringComparison.Ordinal);
                _message = nextMessage;
                _messageTitle.Text = string.IsNullOrWhiteSpace(_message.Subject)
                    ? "(No subject)"
                    : _message.Subject;
                _messageMeta.Text =
                    "From " +
                    (string.IsNullOrWhiteSpace(_message.Sender)
                        ? "unknown sender"
                        : _message.Sender);

                if (changed)
                {
                    _history.Clear();
                    _lastAssistantText = string.Empty;
                    _transcript.Clear();
                    AppendSystem(
                        "Selected email loaded. Ask a question or request draft text.");
                    UpdateDraftButtons();
                }

                SetStatus(
                    "Only this selected email is available to the chat.",
                    false);
            }
            catch (Exception exception)
            {
                _message = null;
                _messageTitle.Text = "No email selected";
                _messageMeta.Text = "Select or open an email, then choose Refresh.";
                SetStatus(exception.Message, true);
                Log.Error("CaptureCurrent", exception);
            }
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(0)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildTranscript(), 0, 2);
            root.Controls.Add(BuildComposer(), 0, 3);
            root.Controls.Add(BuildDraftActions(), 0, 4);
            root.Controls.Add(BuildStatusArea(), 0, 5);
            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted,
                Padding = new Padding(18, 14, 18, 10)
            };

            _messageTitle.AutoEllipsis = true;
            _messageTitle.Dock = DockStyle.Top;
            _messageTitle.Height = 26;
            _messageTitle.Font = new Font(
                Font.FontFamily,
                Font.Size + 3F,
                FontStyle.Bold);
            _messageTitle.ForeColor = TextPrimary;

            _messageMeta.AutoEllipsis = true;
            _messageMeta.Dock = DockStyle.Top;
            _messageMeta.Height = 22;
            _messageMeta.Font = new Font(
                Font.FontFamily,
                Font.Size,
                FontStyle.Regular);
            _messageMeta.ForeColor = TextSecondary;

            header.Controls.Add(_messageMeta);
            header.Controls.Add(_messageTitle);
            return header;
        }

        private Control BuildToolbar()
        {
            var toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(12, 4, 12, 2),
                BackColor = SystemColors.Window
            };

            _refresh = MakeLinkButton("Refresh email", 104);
            _refresh.Click += (sender, args) => RefreshCurrentMessage();
            _settingsButton = MakeLinkButton("Settings", 82);
            _settingsButton.Click += SettingsClick;

            toolbar.Controls.Add(_refresh);
            toolbar.Controls.Add(_settingsButton);
            return toolbar;
        }

        private Control BuildTranscript()
        {
            _transcript.Dock = DockStyle.Fill;
            _transcript.BorderStyle = BorderStyle.None;
            _transcript.BackColor = SystemColors.Window;
            _transcript.ForeColor = TextPrimary;
            _transcript.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _transcript.ReadOnly = true;
            _transcript.DetectUrls = false;
            _transcript.HideSelection = false;
            _transcript.ScrollBars = RichTextBoxScrollBars.Vertical;
            _transcript.Margin = new Padding(18, 8, 18, 8);
            _transcript.AccessibleName = "AI chat conversation";
            _transcript.AccessibleDescription =
                "Plain-text conversation about the selected Outlook email.";

            var frame = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = SystemColors.Window
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
                RowCount = 2,
                Padding = new Padding(18, 8, 18, 8),
                BackColor = SurfaceMuted
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 104));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));

            _composer.Dock = DockStyle.Fill;
            _composer.Multiline = true;
            _composer.AcceptsReturn = true;
            _composer.ScrollBars = ScrollBars.Vertical;
            _composer.Font = new Font(
                Font.FontFamily,
                Font.Size + 1F,
                FontStyle.Regular);
            _composer.BorderStyle = BorderStyle.FixedSingle;
            _composer.MaxLength = TextBoundary.MaxUserPromptCharacters;
            _composer.AccessibleName = "Message to AI";
            _composer.AccessibleDescription =
                "Ask about the selected email or request draft text. Control Enter sends.";

            ConfigurePrimaryButton(_send, "Send to AI");
            _send.Dock = DockStyle.Fill;
            _send.Margin = new Padding(10, 0, 0, 0);
            _send.Click += SendClick;

            var hint = new Label
            {
                AutoSize = true,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text =
                    "Ask about this email or request draft text. " +
                    "Ctrl+Enter sends.",
                Padding = new Padding(0, 3, 0, 0)
            };

            panel.Controls.Add(_composer, 0, 0);
            panel.Controls.Add(_send, 1, 0);
            panel.Controls.Add(hint, 0, 1);
            panel.SetColumnSpan(hint, 2);
            return panel;
        }

        private Control BuildDraftActions()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(18, 3, 18, 3),
                BackColor = SystemColors.Window
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var disclosure = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = TextSecondary,
                Font = new Font(
                    Font.FontFamily,
                    Math.Max(8F, Font.Size - 1F),
                    FontStyle.Regular),
                Text = "Drafts use the entire latest assistant response.",
                AccessibleName = "Draft content disclosure"
            };

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 2, 0, 0),
                BackColor = SystemColors.Window
            };

            ConfigureSecondaryButton(_replyDraft, "Open reply draft", 132);
            _replyDraft.Click += ReplyDraftClick;
            ConfigureSecondaryButton(_newDraft, "Open new draft", 126);
            _newDraft.Click += NewDraftClick;

            actions.Controls.Add(_replyDraft);
            actions.Controls.Add(_newDraft);
            panel.Controls.Add(disclosure, 0, 0);
            panel.Controls.Add(actions, 0, 1);
            return panel;
        }

        private Control BuildStatusArea()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceMuted,
                Padding = new Padding(18, 8, 18, 8)
            };
            _status.Dock = DockStyle.Fill;
            _status.AutoEllipsis = true;
            _status.ForeColor = TextSecondary;
            _status.TextAlign = ContentAlignment.MiddleLeft;
            _status.AccessibleName = "Chat status";
            _status.AccessibleRole = AccessibleRole.StatusBar;
            panel.Controls.Add(_status);
            return panel;
        }

        private async void SendClick(object sender, EventArgs eventArgs)
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
                SetStatus("Type a question or drafting instruction first.", true);
                return;
            }

            if (_message == null)
            {
                SetStatus("Select an email and choose Refresh email.", true);
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

            var requestMessage = _message;
            var requestEntryId = requestMessage.EntryId;
            var requestStoreId = requestMessage.StoreId;
            var transcriptStart = _transcript.TextLength;
            AppendTurn("You", prompt, OutlookBlue);
            _composer.Clear();
            SetBusy(true);
            _requestCancellation = new CancellationTokenSource();

            try
            {
                var response = await _client.CompleteAsync(
                    _settings,
                    requestMessage,
                    _history,
                    prompt,
                    _requestCancellation.Token);

                if (_message == null ||
                    !_message.EntryId.Equals(
                        requestEntryId,
                        StringComparison.Ordinal) ||
                    !_message.StoreId.Equals(
                        requestStoreId,
                        StringComparison.Ordinal))
                {
                    RestoreFailedPrompt(prompt, transcriptStart);
                    SetStatus(
                        "The selected email changed. The response was discarded.",
                        true);
                    return;
                }

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(new ChatTurn("assistant", response));
                _lastAssistantText = response;
                AppendTurn("Assistant", response, TextPrimary);
                SetStatus(
                    "Response received. Draft actions require your click.",
                    false);
            }
            catch (OperationCanceledException)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                SetStatus(
                    _requestCancellation.IsCancellationRequested
                        ? "Request cancelled."
                        : "The AI endpoint timed out. Your prompt was restored.",
                    !_requestCancellation.IsCancellationRequested);
            }
            catch (Exception exception)
            {
                RestoreFailedPrompt(prompt, transcriptStart);
                SetStatus(exception.Message, true);
                Log.Error("CompleteAsync", exception);
            }
            finally
            {
                _requestCancellation.Dispose();
                _requestCancellation = null;
                SetBusy(false);
            }
        }

        private void ReplyDraftClick(object sender, EventArgs eventArgs)
        {
            try
            {
                new DraftService(_outlookApplication)
                    .CreateReplyDraft(_message, _lastAssistantText);
                SetStatus(
                    "Unsent reply draft opened in Outlook for your review.",
                    false);
            }
            catch (Exception exception)
            {
                SetStatus(exception.Message, true);
                Log.Error("CreateReplyDraft", exception);
            }
        }

        private void NewDraftClick(object sender, EventArgs eventArgs)
        {
            try
            {
                new DraftService(_outlookApplication)
                    .CreateNewDraft(_lastAssistantText);
                SetStatus(
                    "Unsent new-message draft opened in Outlook for your review.",
                    false);
            }
            catch (Exception exception)
            {
                SetStatus(exception.Message, true);
                Log.Error("CreateNewDraft", exception);
            }
        }

        private void SettingsClick(object sender, EventArgs eventArgs)
        {
            OpenSettings();
        }

        private void OpenSettings()
        {
            using (var settingsWindow =
                new SettingsWindow(_settingsStore, _settings))
            {
                if (settingsWindow.ShowDialog(this) == DialogResult.OK)
                {
                    _settings = settingsWindow.SavedSettings;
                    SetStatus("AI endpoint settings saved.", false);
                }
            }
        }

        private void AppendSystem(string text)
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Italic);
            _transcript.SelectionColor = TextSecondary;
            _transcript.AppendText(text + Environment.NewLine + Environment.NewLine);
            ScrollTranscript();
        }

        private void AppendTurn(string speaker, string text, Color headingColor)
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            _transcript.SelectionColor = headingColor;
            _transcript.AppendText(speaker + Environment.NewLine);
            _transcript.SelectionFont = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _transcript.SelectionColor = TextPrimary;
            _transcript.AppendText(
                TextBoundary.PlainText(
                    text,
                    speaker == "You"
                        ? TextBoundary.MaxUserPromptCharacters
                        : TextBoundary.MaxAssistantCharacters) +
                Environment.NewLine +
                Environment.NewLine);
            ScrollTranscript();
        }

        private void ScrollTranscript()
        {
            _transcript.SelectionStart = _transcript.TextLength;
            _transcript.ScrollToCaret();
        }

        private void RestoreFailedPrompt(string prompt, int transcriptStart)
        {
            if (_transcript.TextLength > transcriptStart)
            {
                _transcript.Select(
                    transcriptStart,
                    _transcript.TextLength - transcriptStart);
                _transcript.SelectedText = string.Empty;
            }

            _composer.Text = prompt;
            _composer.SelectionStart = _composer.TextLength;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            _send.Text = busy ? "Cancel" : "Send to AI";
            _composer.Enabled = !busy;
            _refresh.Enabled = !busy;
            _settingsButton.Enabled = !busy;
            UpdateDraftButtons();
            if (busy)
            {
                SetStatus("Waiting for the configured AI endpoint...", false);
            }
        }

        private void UpdateDraftButtons()
        {
            var enabled = !_busy && _lastAssistantText.Length > 0;
            _newDraft.Enabled = enabled;
            _replyDraft.Enabled =
                enabled && _message != null && _message.CanReply;
        }

        private void SetStatus(string text, bool error)
        {
            _status.Text = text;
            _status.ForeColor = error
                ? (SystemInformation.HighContrast
                    ? SystemColors.HotTrack
                    : Color.FromArgb(163, 38, 38))
                : TextSecondary;
        }

        private void WindowKeyDown(object sender, KeyEventArgs eventArgs)
        {
            if (eventArgs.Control && eventArgs.KeyCode == Keys.Enter)
            {
                eventArgs.SuppressKeyPress = true;
                SendClick(_send, EventArgs.Empty);
            }
        }

        private void WindowClosed(object sender, FormClosedEventArgs eventArgs)
        {
            _requestCancellation?.Cancel();
            _requestCancellation?.Dispose();
            _client.Dispose();
        }

        private static Button MakeLinkButton(string text, int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Window,
                ForeColor = OutlookBlue,
                UseVisualStyleBackColor = false,
                Margin = new Padding(0, 0, 6, 0),
                AccessibleName = text
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static void ConfigurePrimaryButton(Button button, string text)
        {
            button.Text = text;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = OutlookBlue;
            button.ForeColor = SystemColors.HighlightText;
            button.UseVisualStyleBackColor = false;
            button.FlatAppearance.BorderColor = OutlookBlue;
            button.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Bold);
            button.AccessibleName = text;
        }

        private static void ConfigureSecondaryButton(
            Button button,
            string text,
            int width)
        {
            button.Text = text;
            button.Width = width;
            button.Height = 34;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = SystemColors.Window;
            button.ForeColor = TextPrimary;
            button.UseVisualStyleBackColor = false;
            button.FlatAppearance.BorderColor = Border;
            button.Margin = new Padding(8, 0, 0, 0);
            button.AccessibleName = text;
        }
    }
}
