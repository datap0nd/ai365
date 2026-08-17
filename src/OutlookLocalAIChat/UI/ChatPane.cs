/*
THESIS: A restrained Outlook sidebar makes mailbox retrieval and one linked,
human-reviewed draft visible without granting send capability.
OWN-WORLD: A dark web chat surface rendered by WebView2 from an embedded,
network-isolated page; every piece of model or mailbox text enters the DOM
as inert text nodes, never as HTML. The C# host keeps every capability
boundary; the page is presentation only.
STORY: Ask the mailbox, watch slim activity lines record what loaded, then
deliberately open an unsent draft.
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
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
        private const int MaxTranscriptEvents = 400;
        private const int MaxExternalImages = 4;

        private sealed class ExternalImageContext
        {
            public ExternalImageContext(
                VisionImagePayload payload,
                string thumbnail)
            {
                Payload = payload;
                Thumbnail = thumbnail ?? string.Empty;
            }

            public VisionImagePayload Payload { get; }

            public string Thumbnail { get; }
        }

        private readonly SettingsStore _settingsStore =
            new SettingsStore();
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly List<ChatTurn> _history =
            new List<ChatTurn>();
        private readonly List<MessageSnapshot> _workingMessages =
            new List<MessageSnapshot>();
        private readonly List<ExternalContextDocument> _externalContext =
            new List<ExternalContextDocument>();
        private readonly List<ExternalImageContext> _externalImages =
            new List<ExternalImageContext>();
        private readonly List<string> _transcriptEvents =
            new List<string>();
        private readonly WebView2 _webView = new WebView2();

        private object _outlookApplication;
        private AppSettings _settings;
        private MessageSnapshot _selectedMessage;
        private DraftToolHost _draftTools;
        private CancellationTokenSource _requestCancellation;
        private bool _busy;
        private bool _shutdown;
        private bool _webReady;
        private string _scopeText =
            "No context - use /search or select emails";
        private string _statusText = "Ready";
        private bool _statusError;

        public ChatPane()
        {
            LastCreated = this;
            _settings = _settingsStore.Load();

            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(26, 27, 30);
            MinimumSize = new Size(300, 480);
            AllowDrop = true;
            DragEnter += ChatPaneDragEnter;
            DragDrop += ChatPaneDragDrop;

            _webView.Dock = DockStyle.Fill;
            _webView.DefaultBackgroundColor =
                Color.FromArgb(26, 27, 30);
            Controls.Add(_webView);
            InitializeWebView();
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
        }

        // ------------------------------------------------------------------
        // WebView2 hosting: the embedded page is the only content ever
        // loaded, remote navigation is cancelled, and script has no access
        // to anything but the JSON bridge below.
        // ------------------------------------------------------------------

        private async void InitializeWebView()
        {
            try
            {
                var dataFolder = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder
                            .LocalApplicationData),
                    "MetoAI",
                    "WebView2");
                var environment =
                    await CoreWebView2Environment.CreateAsync(
                        null,
                        dataFolder);
                await _webView.EnsureCoreWebView2Async(environment);
                if (_shutdown)
                {
                    return;
                }

                var settings = _webView.CoreWebView2.Settings;
                settings.AreDefaultContextMenusEnabled = false;
                settings.AreDevToolsEnabled = false;
                settings.IsStatusBarEnabled = false;
                settings.AreBrowserAcceleratorKeysEnabled = false;
                settings.IsBuiltInErrorPageEnabled = false;
                settings.IsZoomControlEnabled = false;
                // External drop stays enabled so the page can receive
                // drag gestures; the page never navigates on drop and
                // forwards the gesture to the host instead.
                _webView.AllowExternalDrop = true;
                _webView.CoreWebView2.NavigationStarting +=
                    WebNavigationStarting;
                _webView.CoreWebView2.NewWindowRequested +=
                    (sender, eventArgs) => eventArgs.Handled = true;
                _webView.CoreWebView2.WebMessageReceived +=
                    WebMessageReceived;
                _webView.CoreWebView2.NavigateToString(
                    LoadChatPage());
            }
            catch (Exception exception)
            {
                Log.Error("WebViewInit", exception);
                ShowWebViewFallback(exception);
            }
        }

        private void ShowWebViewFallback(Exception exception)
        {
            try
            {
                Controls.Remove(_webView);
            }
            catch
            {
            }

            var notice = new Label
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(26, 27, 30),
                ForeColor = Color.FromArgb(232, 232, 236),
                Padding = new Padding(18),
                Text =
                    "MetoAI needs the Microsoft Edge WebView2 " +
                    "runtime, which ships with Windows 10/11 and " +
                    "Microsoft Edge. Install it from Microsoft, then " +
                    "restart Outlook.\r\n\r\nDetails: " +
                    TextBoundary.SingleLine(
                        exception?.Message,
                        300)
            };
            Controls.Add(notice);
        }

        private static string LoadChatPage()
        {
            using (var stream = typeof(ChatPane).Assembly
                .GetManifestResourceStream(
                    "OutlookLocalAIChat.UI.ChatPaneWeb.html"))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "The embedded chat page is missing.");
                }

                using (var reader = new StreamReader(
                    stream,
                    System.Text.Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static void WebNavigationStarting(
            object sender,
            CoreWebView2NavigationStartingEventArgs eventArgs)
        {
            var uri = eventArgs.Uri ?? string.Empty;
            if (!uri.StartsWith(
                    "data:",
                    StringComparison.OrdinalIgnoreCase) &&
                !uri.StartsWith(
                    "about:",
                    StringComparison.OrdinalIgnoreCase))
            {
                eventArgs.Cancel = true;
            }
        }

        private void WebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            try
            {
                var json = eventArgs.TryGetWebMessageAsString();
                var message = _serializer.DeserializeObject(json)
                    as IDictionary<string, object>;
                if (message == null)
                {
                    return;
                }

                object typeValue;
                message.TryGetValue("type", out typeValue);
                var type = Convert.ToString(typeValue) ??
                    string.Empty;
                switch (type)
                {
                    case "ready":
                        HandleWebReady();
                        break;
                    case "send":
                        object textValue;
                        message.TryGetValue("text", out textValue);
                        HandleSendMessage(
                            Convert.ToString(textValue) ??
                            string.Empty);
                        break;
                    case "stop":
                        _requestCancellation?.Cancel();
                        break;
                    case "newChat":
                        HandleNewChat();
                        break;
                    case "addEmail":
                        AddActiveSelection();
                        break;
                    case "addFiles":
                        HandleAddFiles();
                        break;
                    case "openSettings":
                        OpenSettings();
                        break;
                    case "clearContext":
                        HandleClearContext();
                        break;
                    case "removeContext":
                        object kindValue;
                        object indexValue;
                        message.TryGetValue("kind", out kindValue);
                        message.TryGetValue("index", out indexValue);
                        int removeIndex;
                        int.TryParse(
                            Convert.ToString(indexValue),
                            out removeIndex);
                        HandleRemoveContext(
                            Convert.ToString(kindValue) ??
                            string.Empty,
                            removeIndex);
                        break;
                    case "emailDrop":
                        if (!_busy)
                        {
                            AddActiveSelection();
                        }

                        break;
                    case "fileDrop":
                        HandleWebFileDrop(eventArgs);
                        break;
                    case "setModel":
                        object modelValue;
                        message.TryGetValue("model", out modelValue);
                        HandleSetModel(
                            Convert.ToString(modelValue) ??
                            string.Empty);
                        break;
                }
            }
            catch (Exception exception)
            {
                Log.Error("WebMessage", exception);
            }
        }

        private void HandleWebReady()
        {
            _webReady = true;
            RefreshModelPicker();
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "scope" },
                { "text", _scopeText }
            });
            UpdateDraftState();
            PushContextToWeb();
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            foreach (var recorded in _transcriptEvents.ToArray())
            {
                PostRawToWeb(recorded);
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "busy" },
                { "value", _busy }
            });
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "status" },
                { "text", _statusText },
                { "error", _statusError }
            });
        }

        private void PostToWeb(IDictionary<string, object> payload)
        {
            PostRawToWeb(_serializer.Serialize(payload));
        }

        private void PostRawToWeb(string json)
        {
            if (!_webReady || _shutdown)
            {
                return;
            }

            try
            {
                _webView.CoreWebView2?.PostWebMessageAsJson(json);
            }
            catch (Exception exception)
            {
                Log.Error("PostToWeb", exception);
            }
        }

        private void PostTranscript(
            IDictionary<string, object> payload)
        {
            var json = _serializer.Serialize(payload);
            _transcriptEvents.Add(json);
            if (_transcriptEvents.Count > MaxTranscriptEvents)
            {
                _transcriptEvents.RemoveAt(0);
            }

            PostRawToWeb(json);
        }

        // ------------------------------------------------------------------
        // Transcript primitives. Model and mailbox text crosses the bridge
        // as plain strings; the page inserts it via textContent only.
        // ------------------------------------------------------------------

        private void AppendUserTurn(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "user" },
                { "text", text }
            });
        }

        private void AppendFormattedAssistantText(string text)
        {
            var formatted = SafeModelText.Format(
                text,
                TextBoundary.MaxAssistantCharacters);
            var ranges = new List<object>();
            foreach (var range in formatted.BoldRanges)
            {
                ranges.Add(new Dictionary<string, object>
                {
                    { "s", range.Start },
                    { "l", range.Length }
                });
            }

            PostTranscript(new Dictionary<string, object>
            {
                { "type", "assistant" },
                { "text", formatted.PlainText },
                { "bold", ranges }
            });
        }

        private void AppendContext(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.SingleLine(text, 400) },
                { "kind", "context" }
            });
        }

        private void AppendDraftAction(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.SingleLine(text, 400) },
                { "kind", "draft" }
            });
        }

        private void AppendError(string text)
        {
            PostTranscript(new Dictionary<string, object>
            {
                { "type", "activity" },
                { "text", TextBoundary.PlainText(text, 2400) },
                { "kind", "error" }
            });
        }

        private void SetStatus(string text, bool error)
        {
            _statusText = TextBoundary.SingleLine(text, 300);
            _statusError = error;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "status" },
                { "text", _statusText },
                { "error", error }
            });
        }

        private void SetScope(string text)
        {
            _scopeText = TextBoundary.SingleLine(text, 200);
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "scope" },
                { "text", _scopeText }
            });
        }

        private void SetScopeUnavailable(string text)
        {
            SetScope(text);
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "busy" },
                { "value", busy }
            });
            if (busy)
            {
                SetStatus("Thinking...", false);
            }
        }

        private void UpdateDraftState()
        {
            var linked =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "draft" },
                {
                    "text",
                    linked
                        ? "Draft linked - feedback updates it. MetoAI cannot send."
                        : "Say 'create a draft' to open one. MetoAI cannot send."
                },
                { "linked", linked }
            });
        }

        private void RefreshModelPicker()
        {
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

            var items = new List<object>();
            foreach (var model in models)
            {
                if (ModelCatalog.IsDisallowedModel(model))
                {
                    continue;
                }

                items.Add(new Dictionary<string, object>
                {
                    { "id", model },
                    {
                        "vision",
                        ModelCatalog.IsVisionCapable(model)
                    }
                });
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "models" },
                { "items", items },
                { "current", current }
            });
        }

        private void PushContextToWeb()
        {
            var items = new List<object>();
            if (_selectedMessage != null)
            {
                var selectedCard =
                    (Dictionary<string, object>)
                    BuildWorkingSetCard(0, _selectedMessage);
                selectedCard["kind"] = "selected";
                selectedCard["index"] = 0;
                selectedCard["badge"] = "@";
                items.Add(selectedCard);
            }

            for (var index = 0;
                 index < _workingMessages.Count;
                 index++)
            {
                var card =
                    (Dictionary<string, object>)
                    BuildWorkingSetCard(
                        index,
                        _workingMessages[index]);
                card["kind"] = "email";
                card["index"] = index;
                items.Add(card);
            }

            for (var index = 0;
                 index < _externalContext.Count;
                 index++)
            {
                var card =
                    (Dictionary<string, object>)
                    BuildExternalContextCard(
                        _externalContext[index]);
                card["kind"] = "file";
                card["index"] = index;
                items.Add(card);
            }

            for (var index = 0;
                 index < _externalImages.Count;
                 index++)
            {
                var image = _externalImages[index];
                items.Add(new Dictionary<string, object>
                {
                    { "kind", "image" },
                    { "index", index },
                    {
                        "title",
                        TextBoundary.SingleLine(
                            image.Payload.FileName,
                            120)
                    },
                    { "subtitle", "image - vision input" },
                    { "thumb", image.Thumbnail }
                });
            }

            PostToWeb(new Dictionary<string, object>
            {
                { "type", "context" },
                { "items", items }
            });
        }

        private void HandleRemoveContext(string kind, int index)
        {
            if (_busy)
            {
                return;
            }

            switch (kind)
            {
                case "selected":
                    _selectedMessage = null;
                    break;
                case "email":
                    if (index >= 0 &&
                        index < _workingMessages.Count)
                    {
                        _workingMessages.RemoveAt(index);
                    }

                    break;
                case "file":
                    if (index >= 0 &&
                        index < _externalContext.Count)
                    {
                        _externalContext.RemoveAt(index);
                    }

                    break;
                case "image":
                    if (index >= 0 &&
                        index < _externalImages.Count)
                    {
                        _externalImages.RemoveAt(index);
                    }

                    break;
            }

            if (_selectedMessage == null &&
                _workingMessages.Count == 0)
            {
                SetScopeUnavailable(
                    "No context - use /search or select emails");
            }
            else if (_workingMessages.Count > 0)
            {
                SetScope(
                    "Working set: " +
                    _workingMessages.Count +
                    " of " +
                    MailboxWorkingSet.MaxMessages +
                    " emails");
            }

            RefreshContextLayer("External files");
            SetStatus("Removed from context", false);
        }

        private object BuildWorkingSetCard(
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
            return new Dictionary<string, object>
            {
                { "badge", (index + 1).ToString() },
                { "title", subject },
                { "subtitle", sender + "  |  " + date }
            };
        }

        private object BuildExternalContextCard(
            ExternalContextDocument document)
        {
            return new Dictionary<string, object>
            {
                { "badge", "F" },
                {
                    "title",
                    TextBoundary.SingleLine(document.Name, 180)
                },
                {
                    "subtitle",
                    document.Content.Length + " text characters"
                }
            };
        }

        private void RefreshContextLayer(string source)
        {
            PushContextToWeb();
        }

        // ------------------------------------------------------------------
        // Context management (unchanged capability boundaries).
        // ------------------------------------------------------------------

        public void RefreshSelectedMessage()
        {
            if (_outlookApplication == null)
            {
                SetScopeUnavailable(
                    "Outlook is still initializing");
                return;
            }

            if (_busy)
            {
                SetStatus(
                    "Still working - try again in a moment",
                    true);
                return;
            }

            try
            {
                SetSelectedMessage(
                    new MessageReader(_outlookApplication)
                        .CaptureCurrent());
                SetStatus("Email selected", false);
            }
            catch (Exception exception)
            {
                _workingMessages.Clear();
                RefreshContextLayer("External files");
                _selectedMessage = null;
                SetScopeUnavailable(
                    "No context - use /search or select emails");
                SetStatus("Ready", false);
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
                    "Still working - try again in a moment",
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
                SetStatus("Email added to context", false);
                return;
            }

            SetWorkingMessages(
                messages,
                messages.Count + " emails selected in Outlook");
        }

        private void SetSelectedMessage(MessageSnapshot message)
        {
            _workingMessages.Clear();
            RefreshContextLayer("External files");
            _selectedMessage = message ??
                throw new ArgumentNullException(nameof(message));
            var displaySubject = SubjectDisplay.Clean(
                _selectedMessage.Subject);
            SetScope(
                "Selected: " +
                (string.IsNullOrWhiteSpace(displaySubject)
                    ? "(No subject)"
                    : displaySubject));
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
            SetScope(
                "Working set: " +
                _workingMessages.Count +
                " of " +
                MailboxWorkingSet.MaxMessages +
                " emails");
            RefreshContextLayer(source);
            AppendContext(
                TextBoundary.SingleLine(source, 260) +
                " - working set ready");
            SetStatus("Working set ready", false);
        }

        private void HandleClearContext()
        {
            if (_busy)
            {
                return;
            }

            _workingMessages.Clear();
            _externalContext.Clear();
            _externalImages.Clear();
            _selectedMessage = null;
            RefreshContextLayer("External files");
            SetScopeUnavailable(
                "No context - use /search or select emails");
            SetStatus("Context cleared", false);
        }

        private void HandleWebFileDrop(
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            if (_busy)
            {
                return;
            }

            var paths = new List<string>();
            var objects = eventArgs.AdditionalObjects;
            if (objects == null)
            {
                return;
            }

            foreach (var item in objects)
            {
                var file = item as CoreWebView2File;
                if (file != null &&
                    !string.IsNullOrEmpty(file.Path))
                {
                    paths.Add(file.Path);
                }
            }

            if (paths.Count > 0)
            {
                AddExternalFiles(paths);
            }
        }

        private void HandleAddFiles()
        {
            if (_busy)
            {
                return;
            }

            using (var dialog = new OpenFileDialog
            {
                Title = "Add bounded text context to MetoAI",
                Multiselect = true,
                CheckFileExists = true,
                Filter =
                    "Supported files|*.txt;*.md;*.csv;*.json;*.xml;*.html;*.htm;*.log;*.pdf;*.docx;*.pptx;*.xlsx;*.xlsm;*.xls;*.doc;*.ppt;*.rtf;*.eml;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.tif;*.tiff|" +
                    "All files|*.*"
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    AddExternalFiles(dialog.FileNames);
                }
            }
        }

        // Any supported file type is accepted: documents run through the
        // same bounded extractors as email attachments, and images become
        // vision input with a tray thumbnail. ExternalContextLoader
        // remains the strict text-only path for programmatic use.
        private void AddExternalFiles(IEnumerable<string> paths)
        {
            try
            {
                var added = 0;
                foreach (var path in
                    paths ?? new string[0])
                {
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        continue;
                    }

                    var content =
                        EmailAttachmentReader.LoadLocalFile(path);
                    if (content == null)
                    {
                        continue;
                    }

                    if (content.ImageDataUrl.Length > 0)
                    {
                        if (_externalImages.Count >=
                            MaxExternalImages)
                        {
                            SetStatus(
                                "Image limit reached (" +
                                MaxExternalImages + ")",
                                true);
                            continue;
                        }

                        _externalImages.Add(
                            new ExternalImageContext(
                                new VisionImagePayload(
                                    content.FileName,
                                    content.ImageDataUrl),
                                EmailAttachmentReader
                                    .BuildThumbnailDataUrl(path)));
                        AppendContext(
                            "Added image " + content.FileName);
                        added++;
                        continue;
                    }

                    if (content.Text.Length == 0)
                    {
                        continue;
                    }

                    var combined =
                        new List<ExternalContextDocument>(
                            _externalContext);
                    combined.Add(new ExternalContextDocument(
                        content.FileName,
                        content.Text));
                    var normalized =
                        ExternalContextDocument.Normalize(combined);
                    if (normalized.Count <= _externalContext.Count)
                    {
                        SetStatus(
                            "Document limit reached (" +
                            ExternalContextDocument.MaxDocuments +
                            " files, bounded text)",
                            true);
                        continue;
                    }

                    _externalContext.Clear();
                    foreach (var document in normalized)
                    {
                        _externalContext.Add(document);
                    }

                    AppendContext("Added " + content.FileName);
                    added++;
                }

                RefreshContextLayer("External files");
                if (added > 0)
                {
                    SetStatus(
                        added +
                        (added == 1
                            ? " item added"
                            : " items added"),
                        false);
                }
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

        // ------------------------------------------------------------------
        // Local /search command.
        // ------------------------------------------------------------------

        private void HandleLocalSearchCommand(
            string prompt,
            LocalSearchCommand command)
        {
            AppendUserTurn(prompt);
            switch (command.Kind)
            {
                case LocalSearchCommandKind.Help:
                    AppendContext(
                        "/search <person or topic> builds a ten-email " +
                        "working set; /search clear removes it");
                    SetStatus("Search help shown", false);
                    return;
                case LocalSearchCommandKind.Clear:
                    _workingMessages.Clear();
                    RefreshContextLayer("External files");
                    _selectedMessage = null;
                    SetScopeUnavailable(
                        "No context - use /search or select emails");
                    AppendContext("Working set cleared");
                    SetStatus("Working set cleared", false);
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
                    "[OUTLOOK_NOT_READY] Outlook is still initializing",
                    true);
                return;
            }

            SetStatus("Searching mailbox...", false);
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
                        "No emails matched '" + query + "'" +
                        (_workingMessages.Count > 0
                            ? " - previous working set kept"
                            : ""));
                    SetStatus("No matches - refine /search", true);
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
        }

        // ------------------------------------------------------------------
        // Chat request flow.
        // ------------------------------------------------------------------

        private async void HandleSendMessage(string rawText)
        {
            if (_busy)
            {
                return;
            }

            var prompt = TextBoundary.PlainText(
                rawText,
                TextBoundary.MaxUserPromptCharacters);
            if (prompt.Length == 0)
            {
                SetStatus("Type a message first", true);
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
                    "[OUTLOOK_NOT_READY] Outlook is still initializing",
                    true);
                return;
            }

            if (!_settings.IsConfigured)
            {
                OpenSettings();
                if (!_settings.IsConfigured)
                {
                    PostToWeb(new Dictionary<string, object>
                    {
                        { "type", "restorePrompt" },
                        { "text", prompt }
                    });
                    return;
                }
            }

            var requestSelectedMessage = _selectedMessage;
            var requestWorkingMessages =
                new List<MessageSnapshot>(_workingMessages);
            var requestExternalContext =
                new List<ExternalContextDocument>(_externalContext);
            var requestExternalImages =
                new List<VisionImagePayload>();
            foreach (var image in _externalImages)
            {
                requestExternalImages.Add(image.Payload);
            }

            var hasLinkedDraft =
                _draftTools != null &&
                _draftTools.HasActiveDraft;
            var draftAuthorization =
                new OneShotDraftAuthorization(
                    !hasLinkedDraft &&
                    DraftIntentPolicy.AllowsCreate(prompt),
                    hasLinkedDraft &&
                    DraftIntentPolicy.AllowsUpdate(prompt));
            AppendUserTurn(prompt);
            SetBusy(true);
            _requestCancellation =
                new CancellationTokenSource();

            try
            {
                var response = await CompleteMailboxChatAsync(
                    requestSelectedMessage,
                    requestWorkingMessages,
                    requestExternalContext,
                    requestExternalImages,
                    prompt,
                    draftAuthorization,
                    _requestCancellation.Token);

                _history.Add(new ChatTurn("user", prompt));
                _history.Add(
                    new ChatTurn("assistant", response));
                AppendFormattedAssistantText(response);
                if (draftAuthorization.IsCreated)
                {
                    SetStatus(
                        "Draft created - unsent, open for review",
                        false);
                }
                else if (draftAuthorization.IsUpdated)
                {
                    SetStatus("Draft updated", false);
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
                    SetStatus("No draft was created", false);
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
                PostToWeb(new Dictionary<string, object>
                {
                    { "type", "restorePrompt" },
                    { "text", prompt }
                });
                SetStatus("Stopped - prompt restored", false);
            }
            catch (Exception exception)
            {
                var details = DiagnosticDetails.ForException(
                    exception,
                    "AI_REQUEST_FAILED");
                AppendError(details);
                PostToWeb(new Dictionary<string, object>
                {
                    { "type", "restorePrompt" },
                    { "text", prompt }
                });
                SetStatus(FirstLine(details), true);
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
            IReadOnlyList<VisionImagePayload> externalImages,
            string prompt,
            OneShotDraftAuthorization draftAuthorization,
            CancellationToken cancellationToken)
        {
            var activeDraft = draftAuthorization.CanUpdate
                ? _draftTools?.ActiveDraft
                : null;
            var imagesExpected = ModelRouting.ContextMayIncludeImages(
                selectedMessage,
                workingMessages) ||
                externalImages.Count > 0;
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
                    activeModel + " is text-only - images will " +
                    "not be read",
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
                SetStatus("Images attached for vision", false);
            }

            if (externalImages.Count > 0)
            {
                VisionAttachmentExchange.AppendVisionContext(
                    request,
                    activeModel,
                    new[]
                    {
                        new MailboxToolResult(
                            "external_files",
                            string.Empty,
                            string.Empty,
                            externalImages)
                    });
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
                        AppendDraftAction(result.StatusText);
                    }
                    else
                    {
                        AppendContext(result.StatusText);
                    }

                    SetStatus(result.StatusText, false);
                }

                activeModel = ModelRouting.ResolveForRequest(
                    _settings,
                    imagesExpected,
                    results);
                request.model = TextBoundary.PlainText(
                    activeModel,
                    200);
                if (ModelRouting.IsTemporaryVisionSwitch(
                        _settings,
                        activeModel))
                {
                    SetStatus(
                        "Using " + activeModel + " for images",
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

        // ------------------------------------------------------------------
        // Chat lifecycle and settings.
        // ------------------------------------------------------------------

        private void HandleNewChat()
        {
            if (_busy)
            {
                return;
            }

            _history.Clear();
            _workingMessages.Clear();
            _externalContext.Clear();
            _externalImages.Clear();
            _transcriptEvents.Clear();
            _draftTools?.Dispose();
            _draftTools = _outlookApplication == null
                ? null
                : new DraftToolHost(_outlookApplication);
            PostToWeb(new Dictionary<string, object>
            {
                { "type", "clear" }
            });
            PushContextToWeb();
            RefreshSelectedMessage();
            UpdateDraftState();
            SetStatus("Chat and context cleared", false);
        }

        private void HandleSetModel(string model)
        {
            if (_busy || model.Length == 0)
            {
                return;
            }

            if (string.Equals(
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
                SetStatus("Model: " + model, false);
            }
            catch (Exception exception)
            {
                Log.Error("SwitchModel", exception);
                SetStatus("The model change was not saved", true);
            }
        }

        private void OpenSettings()
        {
            if (_busy)
            {
                return;
            }

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
            try
            {
                _webView.Dispose();
            }
            catch (Exception exception)
            {
                Log.Error("WebViewDispose", exception);
            }

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

        private static string FirstLine(string value)
        {
            var text = value ?? string.Empty;
            var index = text.IndexOfAny(
                new[] { '\r', '\n' });
            return index >= 0
                ? text.Substring(0, index)
                : text;
        }
    }
}
