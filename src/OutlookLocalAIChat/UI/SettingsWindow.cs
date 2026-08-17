using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.UI
{
    public sealed class SettingsWindow : Form
    {
        private const int ModelDiscoveryTimeoutSeconds = 15;
        private const int EndpointProbeTimeoutSeconds = 90;

        private readonly TextBox _endpoint = new TextBox();
        private readonly ComboBox _model = new ComboBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly CheckBox _allowInsecureHttp = new CheckBox();
        private readonly Label _transportWarning = new Label();
        private readonly CheckBox _switchVisionForImages = new CheckBox();
        private readonly Label _modelGuidance = new Label();
        private readonly ListBox _modelGuide = new ListBox();
        private readonly Label _testStatus = new Label();
        private readonly CheckBox _useToneProfile = new CheckBox();
        private readonly RichTextBox _toneProfile = new RichTextBox();
        private readonly Label _toneStatus = new Label();
        private readonly Label _error = new Label();
        private readonly Button _checkEndpoint =
            MakeButton("Check endpoint", false, 128);
        private readonly Button _refreshModels =
            MakeButton("Refresh models", false, 120);
        private readonly Button _analyzeTone =
            MakeButton("Analyze 15 sent emails", false, 176);
        private readonly Button _save =
            MakeButton("Save", true, 96);
        private readonly OpenAiCompatibleClient _client =
            new OpenAiCompatibleClient();
        private readonly SettingsStore _store;
        private readonly object _outlookApplication;
        private CancellationTokenSource _checkCancellation;
        private CancellationTokenSource _toneCancellation;
        private bool _checking;
        private bool _analyzingTone;
        private bool _refreshingModels;

        public SettingsWindow(
            SettingsStore store,
            AppSettings current)
            : this(store, current, null)
        {
        }

        public SettingsWindow(
            SettingsStore store,
            AppSettings current,
            object outlookApplication)
        {
            _store = store ??
                throw new ArgumentNullException(nameof(store));
            _outlookApplication = outlookApplication;

            Text = "MailAI settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(700, 760);
            MinimumSize = new Size(620, 700);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
            AutoScaleMode = AutoScaleMode.Font;

            ConfigureFields();
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18, 16, 18, 14)
            };
            root.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 46));

            var tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                AccessibleName = "MailAI settings sections"
            };
            tabs.TabPages.Add(BuildConnectionPage());
            tabs.TabPages.Add(BuildWritingStylePage());
            root.Controls.Add(tabs, 0, 0);

            ConfigureSupportingLabel(_error);
            _error.ForeColor = ErrorText;
            _error.AccessibleName = "Settings error";
            _error.AccessibleRole = AccessibleRole.Alert;
            root.Controls.Add(_error, 0, 1);
            var buttons = BuildButtons();
            root.Controls.Add(buttons, 0, 2);
            Controls.Add(root);
            FormClosing += SettingsWindowFormClosing;

            AcceptButton = _save;
            CancelButton = GetCancelButton(buttons);

            _endpoint.Text = current?.BaseUrl ?? string.Empty;
            _model.Text = current?.Model ?? string.Empty;
            _apiKey.Text = current?.ApiKey ?? string.Empty;
            _allowInsecureHttp.Checked =
                current?.AllowInsecureHttp ?? false;
            _toneProfile.Text = TextBoundary.PlainText(
                current?.ToneProfile,
                TextBoundary.MaxToneProfileCharacters);
            _useToneProfile.Checked =
                (current?.UseToneProfile ?? false) &&
                _toneProfile.TextLength > 0;
            _switchVisionForImages.Checked =
                current?.SwitchToVisionModelForImages ?? false;
            RestoreDiscoveredModels(current?.DiscoveredModels);
            UpdateModelGuidance();
            UpdateTransportWarning();
        }

        public AppSettings SavedSettings { get; private set; }

        protected override void OnFormClosed(
            FormClosedEventArgs eventArgs)
        {
            _checkCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _checkCancellation?.Dispose();
            _toneCancellation?.Dispose();
            _client.Dispose();
            base.OnFormClosed(eventArgs);
        }

        private static Color SecondaryText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.GrayText
                    : Color.FromArgb(80, 80, 80);
            }
        }

        private static Color ErrorText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.HotTrack
                    : Color.FromArgb(163, 38, 38);
            }
        }

        private static Color SuccessText
        {
            get
            {
                return SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(20, 112, 70);
            }
        }

        private void ConfigureFields()
        {
            ConfigureField(
                _endpoint,
                "AI endpoint",
                "HTTPS endpoint, loopback HTTP, or explicitly allowed remote HTTP.");
            ConfigureModelField();
            ConfigureField(
                _apiKey,
                "API key",
                "Encrypted for the current Windows user.");
            _apiKey.UseSystemPasswordChar = true;

            _allowInsecureHttp.AutoSize = true;
            _allowInsecureHttp.Text =
                "Allow insecure HTTP for non-local endpoints";
            _allowInsecureHttp.AccessibleName =
                "Allow insecure HTTP";
            _allowInsecureHttp.AccessibleDescription =
                "Allows the API key, prompts, and email context to be sent " +
                "without transport encryption.";
            _allowInsecureHttp.CheckedChanged += InsecureHttpChanged;

            _switchVisionForImages.AutoSize = true;
            _switchVisionForImages.Text =
                "Temporarily switch to the best available vision model when email images are involved";
            _switchVisionForImages.AccessibleName =
                "Switch to vision model for images";
            _switchVisionForImages.AccessibleDescription =
                "Uses your saved model list to pick a vision model for this request only. " +
                "Save settings after Refresh models so MailAI knows which vision models are available.";

            _useToneProfile.AutoSize = true;
            _useToneProfile.Text =
                "Use this writing profile for drafts";
            _useToneProfile.AccessibleDescription =
                "Applies the editable writing profile only when creating or updating drafts.";

            _toneProfile.Dock = DockStyle.Fill;
            _toneProfile.BorderStyle = BorderStyle.FixedSingle;
            _toneProfile.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _toneProfile.MaxLength =
                TextBoundary.MaxToneProfileCharacters;
            _toneProfile.ScrollBars =
                RichTextBoxScrollBars.Vertical;
            _toneProfile.DetectUrls = false;
            _toneProfile.AccessibleName =
                "Editable email writing profile";
        }

        private TabPage BuildConnectionPage()
        {
            var page = new TabPage("Connection");
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 15,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            layout.Controls.Add(FieldLabel("Endpoint or base URL"), 0, 0);
            layout.Controls.Add(_endpoint, 0, 1);
            layout.Controls.Add(FieldLabel("Model"), 0, 2);
            layout.Controls.Add(_model, 0, 3);
            ConfigureSupportingLabel(_modelGuidance);
            layout.Controls.Add(_modelGuidance, 0, 4);
            layout.Controls.Add(FieldLabel("Model guide"), 0, 5);
            ConfigureModelGuide();
            layout.Controls.Add(_modelGuide, 0, 6);
            layout.Controls.Add(_switchVisionForImages, 0, 7);
            layout.Controls.Add(FieldLabel("API key"), 0, 8);
            layout.Controls.Add(_apiKey, 0, 9);
            layout.Controls.Add(_allowInsecureHttp, 0, 10);
            ConfigureSupportingLabel(_transportWarning);
            _transportWarning.AccessibleRole = AccessibleRole.Alert;
            layout.Controls.Add(_transportWarning, 0, 11);

            var hint = SupportingText(
                "The endpoint receives your prompt, conversation, and only bounded " +
                "context. MailAI exposes read tools and guarded draft creation only. " +
                "It has no send, move, delete, or mailbox mutation capability.");
            layout.Controls.Add(hint, 0, 12);
            ConfigureSupportingLabel(_testStatus);
            _testStatus.Text =
                "Use Refresh models to load the model list from your endpoint. " +
                "Check endpoint also verifies tool-call compatibility.";
            _testStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_testStatus, 0, 13);

            _checkEndpoint.Click += CheckEndpointClick;
            _refreshModels.Click += RefreshModelsClick;
            var checkRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0)
            };
            checkRow.Controls.Add(_checkEndpoint);
            checkRow.Controls.Add(_refreshModels);
            layout.Controls.Add(checkRow, 0, 14);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildWritingStylePage()
        {
            var page = new TabPage("Writing style");
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 8,
                Padding = new Padding(18, 16, 18, 12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var disclosure = SupportingText(
                "Nothing is analyzed automatically. When you click Analyze, MailAI " +
                "reads up to 15 recent Sent Items messages, removes obvious quoted " +
                "history, and sends bounded samples to your configured endpoint. " +
                "Review and edit the result before saving.");
            disclosure.ForeColor = SystemColors.ControlText;
            layout.Controls.Add(disclosure, 0, 0);
            layout.Controls.Add(_useToneProfile, 0, 1);
            layout.Controls.Add(FieldLabel("My drafting instructions"), 0, 2);
            layout.Controls.Add(_toneProfile, 0, 3);

            var scope = SupportingText(
                "The profile affects wording, greeting, cadence, and sign-off only. " +
                "It cannot change MailAI permissions or security rules.");
            layout.Controls.Add(scope, 0, 4);

            _analyzeTone.Click += AnalyzeToneClick;
            var actionRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0, 8, 0, 0)
            };
            actionRow.Controls.Add(_analyzeTone);
            layout.Controls.Add(actionRow, 0, 5);

            ConfigureSupportingLabel(_toneStatus);
            _toneStatus.Text =
                "Analysis requires at least five usable sent messages and never runs without this button.";
            _toneStatus.AccessibleRole = AccessibleRole.StatusBar;
            layout.Controls.Add(_toneStatus, 0, 6);
            layout.Controls.Add(
                SupportingText(
                    "Tip: keep the profile general. Remove names, client details, " +
                    "project facts, and anything you would not want reused."),
                0,
                7);
            page.Controls.Add(layout);
            return page;
        }

        private void ConfigureModelGuide()
        {
            _modelGuide.Dock = DockStyle.Fill;
            _modelGuide.IntegralHeight = false;
            _modelGuide.BorderStyle = BorderStyle.FixedSingle;
            _modelGuide.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size,
                FontStyle.Regular);
            _modelGuide.AccessibleName = "Model guide";
            _modelGuide.AccessibleDescription =
                "Reference list of common local models with speed, quality, and vision notes. " +
                "Double-click an entry to pick a matching discovered model.";
            foreach (var entry in ModelCatalog.GuideEntries)
            {
                _modelGuide.Items.Add(entry);
            }

            _modelGuide.DisplayMember = "SummaryLine";
            _modelGuide.SelectedIndexChanged += ModelGuideSelectionChanged;
            _modelGuide.DoubleClick += ModelGuideDoubleClick;
        }

        private void ModelGuideSelectionChanged(
            object sender,
            EventArgs eventArgs)
        {
            var entry = _modelGuide.SelectedItem as ModelGuideEntry;
            if (entry == null)
            {
                return;
            }

            _modelGuidance.Text = entry.SummaryLine +
                "\n" +
                entry.Notes;
        }

        private void ModelGuideDoubleClick(
            object sender,
            EventArgs eventArgs)
        {
            var entry = _modelGuide.SelectedItem as ModelGuideEntry;
            if (entry == null)
            {
                return;
            }

            var discovered = ModelCatalog.FindMatchingDiscoveredId(
                entry,
                _model.Items.Cast<object>()
                    .Select(item => item.ToString()));
            if (discovered.Length > 0)
            {
                _model.Text = discovered;
                UpdateModelGuidance();
            }
        }

        private void ConfigureModelField()
        {
            _model.Dock = DockStyle.Fill;
            _model.DropDownStyle = ComboBoxStyle.DropDown;
            _model.FlatStyle = FlatStyle.Flat;
            _model.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            _model.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _model.AutoCompleteSource = AutoCompleteSource.ListItems;
            _model.MaxDropDownItems = 12;
            _model.AccessibleName = "AI model";
            _model.TextChanged +=
                (sender, args) => UpdateModelGuidance();
        }

        private Control BuildButtons()
        {
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            _save.Click += SaveClick;
            var cancel = MakeButton("Cancel", false, 96);
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Name = "CancelSettings";
            panel.Controls.Add(_save);
            panel.Controls.Add(cancel);
            return panel;
        }

        private static Button GetCancelButton(Control root)
        {
            var matches = root.Controls.Find("CancelSettings", true);
            return matches.Length > 0 ? (Button)matches[0] : null;
        }

        private static void ConfigureField(
            TextBox field,
            string accessibleName,
            string accessibleDescription)
        {
            field.Dock = DockStyle.Fill;
            field.BorderStyle = BorderStyle.FixedSingle;
            field.Font = new Font(
                SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.Size + 1F,
                FontStyle.Regular);
            field.AccessibleName = accessibleName;
            field.AccessibleDescription = accessibleDescription;
        }

        private static void ConfigureSupportingLabel(Label label)
        {
            label.AutoSize = true;
            label.MaximumSize = new Size(620, 0);
            label.ForeColor = SecondaryText;
        }

        private static Label SupportingText(string text)
        {
            var label = new Label { Text = text };
            ConfigureSupportingLabel(label);
            return label;
        }

        private static Label FieldLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Font = new Font(
                    SystemFonts.MessageBoxFont.FontFamily,
                    SystemFonts.MessageBoxFont.Size,
                    FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                Padding = new Padding(0, 0, 0, 4),
                Text = text
            };
        }

        private static Button MakeButton(
            string text,
            bool primary,
            int width)
        {
            var button = new Button
            {
                Text = text,
                Width = width,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(8, 0, 0, 0),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = primary
                ? (SystemInformation.HighContrast
                    ? SystemColors.Highlight
                    : Color.FromArgb(0, 95, 184))
                : SystemColors.Window;
            button.ForeColor = primary
                ? SystemColors.HighlightText
                : SystemColors.WindowText;
            button.FlatAppearance.BorderColor = primary
                ? button.BackColor
                : SystemColors.ControlDark;
            button.AccessibleName = text;
            return button;
        }

        private AppSettings ReadFormSettings()
        {
            var profile = TextBoundary.PlainText(
                _toneProfile.Text,
                TextBoundary.MaxToneProfileCharacters);
            return new AppSettings
            {
                BaseUrl = _endpoint.Text,
                Model = _model.Text,
                ApiKey = _apiKey.Text,
                AllowInsecureHttp = _allowInsecureHttp.Checked,
                ToneProfile = profile,
                UseToneProfile =
                    _useToneProfile.Checked && profile.Length > 0,
                SwitchToVisionModelForImages =
                    _switchVisionForImages.Checked,
                DiscoveredModels = CollectDiscoveredModels()
            };
        }

        private List<string> CollectDiscoveredModels()
        {
            return _model.Items
                .Cast<object>()
                .Select(item =>
                    TextBoundary.PlainText(
                        Convert.ToString(item),
                        200))
                .Where(item => item.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void RestoreDiscoveredModels(
            IEnumerable<string> models)
        {
            foreach (var model in models ?? Enumerable.Empty<string>())
            {
                if (ModelCatalog.IsDisallowedModel(model))
                {
                    continue;
                }

                if (!_model.Items.Cast<object>().Any(item =>
                    string.Equals(
                        item.ToString(),
                        model,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    _model.Items.Add(model);
                }
            }
        }

        private async void AnalyzeToneClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_analyzingTone)
            {
                _toneCancellation?.Cancel();
                return;
            }

            _error.Text = string.Empty;
            if (_outlookApplication == null)
            {
                _error.Text =
                    "[OUTLOOK_NOT_READY] Open MailAI from Outlook before analyzing your writing style.";
                return;
            }

            var settings = ReadFormSettings();
            if (!settings.IsConfigured)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Configure the endpoint, model, and API key first.";
                return;
            }

            SetToneAnalyzing(true);
            _toneCancellation = new CancellationTokenSource();
            try
            {
                _toneStatus.Text =
                    "Reading and cleaning up to 15 recent Sent Items messages...";
                var samples = new SentMailToneSampler(
                    _outlookApplication).CaptureRecent();
                if (samples.Count < 5)
                {
                    throw new AiEndpointException(
                        "TONE_SAMPLES_INSUFFICIENT",
                        "At least five usable sent messages are required. Found " +
                        samples.Count + ".");
                }

                _toneStatus.Text =
                    "Analyzing " + samples.Count +
                    " bounded sent-email samples with " + settings.Model + "...";
                var request = ToneProfileRequestFactory.Create(
                    settings.Model,
                    samples);
                var response = await _client.CompleteAsync(
                    settings,
                    request,
                    _toneCancellation.Token);
                if (response.tool_calls != null &&
                    response.tool_calls.Count > 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_INVALID",
                        "The model returned tool calls during style analysis.");
                }

                var profile = SafeModelText.Format(
                    response.content,
                    TextBoundary.MaxToneProfileCharacters).PlainText;
                if (profile.Length == 0)
                {
                    throw new AiEndpointException(
                        "TONE_RESPONSE_EMPTY",
                        "The model returned an empty writing profile.");
                }

                _toneProfile.Text = profile;
                _useToneProfile.Checked = true;
                _toneStatus.ForeColor = SuccessText;
                _toneStatus.Text =
                    "Writing profile generated from " + samples.Count +
                    " sent messages. Review and edit it, then click Save.";
            }
            catch (OperationCanceledException)
            {
                _toneStatus.Text =
                    "Writing-style analysis was cancelled.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "TONE_ANALYSIS_FAILED");
            }
            finally
            {
                _toneCancellation?.Dispose();
                _toneCancellation = null;
                SetToneAnalyzing(false);
            }
        }

        private async void CheckEndpointClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_checking)
            {
                _checkCancellation?.Cancel();
                return;
            }

            _error.Text = string.Empty;
            var settings = ReadFormSettings();
            if (!settings.HasConnectionSettings)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint and API key first.";
                return;
            }

            SetChecking(true);
            _checkCancellation = new CancellationTokenSource();
            try
            {
                IReadOnlyList<string> models = null;
                var discoveryNote = string.Empty;
                try
                {
                    _testStatus.Text =
                        "Checking authentication and available models " +
                        "(up to " + ModelDiscoveryTimeoutSeconds +
                        " seconds)...";
                    using (var modelsTimeout =
                        CancellationTokenSource.CreateLinkedTokenSource(
                            _checkCancellation.Token))
                    {
                        modelsTimeout.CancelAfter(
                            TimeSpan.FromSeconds(
                                ModelDiscoveryTimeoutSeconds));
                        models = await _client.GetModelsAsync(
                            settings,
                            modelsTimeout.Token);
                    }

                    AddDiscoveredModels(models);
                }
                catch (OperationCanceledException)
                {
                    if (_checkCancellation.IsCancellationRequested)
                    {
                        throw;
                    }

                    discoveryNote =
                        " Model discovery timed out after " +
                        ModelDiscoveryTimeoutSeconds +
                        " seconds, so the entered model was tested directly.";
                }
                catch (AiEndpointException exception)
                {
                    discoveryNote =
                        " Model discovery was unavailable [" + exception.Code +
                        "], so the entered model was tested directly.";
                }

                settings = ReadFormSettings();
                if (settings.Model.Trim().Length == 0)
                {
                    throw new AiEndpointException(
                        "MODEL_REQUIRED",
                        "Choose a model or use Refresh models after model discovery returns at least one generative model.");
                }

                _testStatus.Text =
                    "Testing mailbox tool calls with " + settings.Model +
                    " (up to " + EndpointProbeTimeoutSeconds +
                    " seconds)...";
                using (var probeTimeout =
                    CancellationTokenSource.CreateLinkedTokenSource(
                        _checkCancellation.Token))
                {
                    probeTimeout.CancelAfter(
                        TimeSpan.FromSeconds(
                            EndpointProbeTimeoutSeconds));
                    var probe = ChatRequestFactory.CreateEndpointCheck(
                        settings.Model);
                    var response = await _client.CompleteAsync(
                        settings,
                        probe,
                        probeTimeout.Token);
                    var validCall = response.tool_calls != null &&
                        response.tool_calls.Any(call =>
                            call?.function != null &&
                            MailboxToolCatalog.IsApproved(
                                call.function.name));
                    if (!validCall)
                    {
                        throw new AiEndpointException(
                            "MODEL_TOOL_CALL_UNSUPPORTED",
                            "The endpoint answered, but this model did not return a compatible mailbox tool call.");
                    }
                }

                _testStatus.ForeColor = SuccessText;
                _testStatus.Text =
                    "Endpoint verified. Authentication, model, and mailbox tool calling passed." +
                    discoveryNote;
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[ENDPOINT_CHECK_CANCELLED] The endpoint check was cancelled or timed out. " +
                    "Model discovery allows up to " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds. The tool-call probe allows up to " +
                    EndpointProbeTimeoutSeconds +
                    " seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "ENDPOINT_CHECK_FAILED");
            }
            finally
            {
                _checkCancellation?.Dispose();
                _checkCancellation = null;
                SetChecking(false);
            }
        }

        private async void RefreshModelsClick(
            object sender,
            EventArgs eventArgs)
        {
            if (_refreshingModels || _checking)
            {
                return;
            }

            _error.Text = string.Empty;
            var settings = ReadFormSettings();
            if (!settings.HasConnectionSettings)
            {
                _error.Text =
                    "[CONFIGURATION_INCOMPLETE] Enter a valid endpoint and API key first.";
                return;
            }

            _refreshingModels = true;
            _refreshModels.Enabled = false;
            try
            {
                _testStatus.ForeColor = SecondaryText;
                _testStatus.Text =
                    "Loading models from the endpoint (up to " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds)...";
                using (var cancellation = new CancellationTokenSource(
                    TimeSpan.FromSeconds(ModelDiscoveryTimeoutSeconds)))
                {
                    var models = await _client.GetModelsAsync(
                        settings,
                        cancellation.Token);
                    var added = AddDiscoveredModels(models);
                    _testStatus.ForeColor = SuccessText;
                    _testStatus.Text =
                        "Loaded " +
                        added.ToString() +
                        " generative models into the model list. " +
                        "Gauss and Gausso models are excluded automatically.";
                }
            }
            catch (OperationCanceledException)
            {
                _error.Text =
                    "[MODEL_REFRESH_TIMEOUT] Model discovery timed out after " +
                    ModelDiscoveryTimeoutSeconds +
                    " seconds.";
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "MODEL_REFRESH_FAILED");
            }
            finally
            {
                _refreshingModels = false;
                _refreshModels.Enabled =
                    !_checking && !_analyzingTone;
            }
        }

        private int AddDiscoveredModels(IEnumerable<string> models)
        {
            var added = 0;
            foreach (var model in models ?? Enumerable.Empty<string>())
            {
                if (ModelCatalog.IsDisallowedModel(model))
                {
                    continue;
                }

                if (!_model.Items.Cast<object>().Any(item =>
                    string.Equals(
                        item.ToString(),
                        model,
                        StringComparison.OrdinalIgnoreCase)))
                {
                    _model.Items.Add(model);
                    added++;
                }
            }

            if (added > 0 && _model.Text.Trim().Length == 0)
            {
                _modelGuide.SelectedIndex = -1;
            }

            UpdateModelGuidance();
            return added;
        }

        private void SetChecking(bool checking)
        {
            _checking = checking;
            _checkEndpoint.Text =
                checking ? "Cancel check" : "Check endpoint";
            SetCommonControlsEnabled(!checking && !_analyzingTone);
            if (checking)
            {
                _testStatus.ForeColor = SecondaryText;
            }
        }

        private void SetToneAnalyzing(bool analyzing)
        {
            _analyzingTone = analyzing;
            _analyzeTone.Text =
                analyzing ? "Cancel analysis" : "Analyze 15 sent emails";
            SetCommonControlsEnabled(!analyzing && !_checking);
            _analyzeTone.Enabled = analyzing || !_checking;
            if (analyzing)
            {
                _toneStatus.ForeColor = SecondaryText;
            }
        }

        private void SetCommonControlsEnabled(bool enabled)
        {
            _save.Enabled = enabled;
            _endpoint.Enabled = enabled;
            _model.Enabled = enabled;
            _apiKey.Enabled = enabled;
            _allowInsecureHttp.Enabled = enabled;
            _switchVisionForImages.Enabled = enabled;
            _useToneProfile.Enabled = enabled;
            _toneProfile.Enabled = enabled;
            _checkEndpoint.Enabled = enabled || _checking;
            _refreshModels.Enabled =
                (enabled || _refreshingModels) &&
                !_checking &&
                !_refreshingModels;
            _analyzeTone.Enabled = enabled || _analyzingTone;
        }

        private void SaveClick(object sender, EventArgs eventArgs)
        {
            try
            {
                _error.Text = string.Empty;
                var settings = ReadFormSettings();
                _store.Save(settings);
                SavedSettings = settings;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                _error.Text = DiagnosticDetails.ForException(
                    exception,
                    "SETTINGS_SAVE_FAILED");
            }
        }

        private void SettingsWindowFormClosing(
            object sender,
            FormClosingEventArgs eventArgs)
        {
            if (!_checking && !_analyzingTone)
            {
                return;
            }

            eventArgs.Cancel = true;
            _checkCancellation?.Cancel();
            _toneCancellation?.Cancel();
            _error.Text =
                "Cancelling the active settings operation. Close again when it finishes.";
        }

        private void InsecureHttpChanged(object sender, EventArgs eventArgs)
        {
            UpdateTransportWarning();
        }

        private void UpdateModelGuidance()
        {
            var text = _model.Text.Trim();
            if (text.Length == 0)
            {
                _modelGuidance.Text = ModelCatalog.BuildGuideOverview();
                _modelGuide.ClearSelected();
                return;
            }

            _modelGuidance.Text =
                ModelSelectionPolicy.DescriptionFor(text);

            var profile = ModelCatalog.Resolve(text);
            if (profile == null)
            {
                _modelGuide.ClearSelected();
                return;
            }

            for (var index = 0; index < _modelGuide.Items.Count; index++)
            {
                if (_modelGuide.Items[index] == profile)
                {
                    _modelGuide.SelectedIndex = index;
                    return;
                }
            }

            _modelGuide.ClearSelected();
        }

        private void UpdateTransportWarning()
        {
            if (_allowInsecureHttp.Checked)
            {
                _transportWarning.ForeColor = ErrorText;
                _transportWarning.Text =
                    "Warning: with HTTP, the API key, prompts, and retrieved email " +
                    "context cross the network without transport encryption.";
                return;
            }

            _transportWarning.ForeColor = SecondaryText;
            _transportWarning.Text =
                "Loopback HTTP remains available without this setting. " +
                "HTTPS is recommended for every remote endpoint.";
        }
    }
}
