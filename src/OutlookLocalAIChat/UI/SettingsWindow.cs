using System;
using System.Drawing;
using System.Windows.Forms;
using OutlookLocalAIChat.Configuration;

namespace OutlookLocalAIChat.UI
{
    public sealed class SettingsWindow : Form
    {
        private readonly TextBox _endpoint = new TextBox();
        private readonly TextBox _model = new TextBox();
        private readonly TextBox _apiKey = new TextBox();
        private readonly Label _error = new Label();
        private readonly SettingsStore _store;

        public SettingsWindow(SettingsStore store, AppSettings current)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));

            Text = "AI endpoint settings";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(520, 365);
            MinimumSize = new Size(480, 390);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            BackColor = Color.FromArgb(250, 250, 250);
            AutoScaleMode = AutoScaleMode.Dpi;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 10,
                Padding = new Padding(24, 20, 24, 20)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));

            ConfigureField(_endpoint);
            ConfigureField(_model);
            ConfigureField(_apiKey);
            _apiKey.UseSystemPasswordChar = true;

            layout.Controls.Add(FieldLabel("Endpoint or base URL"), 0, 0);
            layout.Controls.Add(_endpoint, 0, 1);
            layout.Controls.Add(FieldLabel("Model"), 0, 2);
            layout.Controls.Add(_model, 0, 3);
            layout.Controls.Add(FieldLabel("API key"), 0, 4);
            layout.Controls.Add(_apiKey, 0, 5);

            var hint = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(460, 0),
                ForeColor = Color.FromArgb(80, 80, 80),
                Text =
                    "Use HTTPS, or HTTP only for a local endpoint. " +
                    "The key is encrypted for your Windows account."
            };
            layout.Controls.Add(hint, 0, 6);

            _error.AutoSize = true;
            _error.ForeColor = Color.FromArgb(163, 38, 38);
            _error.Padding = new Padding(0, 10, 0, 0);
            layout.Controls.Add(_error, 0, 7);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 5, 0, 0)
            };
            var save = MakeButton("Save", true);
            save.Click += SaveClick;
            var cancel = MakeButton("Cancel", false);
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);
            layout.Controls.Add(buttons, 0, 9);

            Controls.Add(layout);
            AcceptButton = save;
            CancelButton = cancel;

            _endpoint.Text = current?.BaseUrl ?? string.Empty;
            _model.Text = current?.Model ?? string.Empty;
            _apiKey.Text = current?.ApiKey ?? string.Empty;
        }

        public AppSettings SavedSettings { get; private set; }

        private static void ConfigureField(TextBox field)
        {
            field.Dock = DockStyle.Fill;
            field.BorderStyle = BorderStyle.FixedSingle;
            field.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
        }

        private static Label FieldLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(35, 35, 35),
                Padding = new Padding(0, 0, 0, 4),
                Text = text
            };
        }

        private static Button MakeButton(string text, bool primary)
        {
            var button = new Button
            {
                Text = text,
                Width = 96,
                Height = 32,
                FlatStyle = FlatStyle.Flat,
                Margin = new Padding(8, 0, 0, 0),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = primary
                ? Color.FromArgb(0, 95, 184)
                : Color.White;
            button.ForeColor = primary ? Color.White : Color.FromArgb(35, 35, 35);
            button.FlatAppearance.BorderColor = primary
                ? Color.FromArgb(0, 95, 184)
                : Color.FromArgb(170, 170, 170);
            return button;
        }

        private void SaveClick(object sender, EventArgs eventArgs)
        {
            try
            {
                var settings = new AppSettings
                {
                    BaseUrl = _endpoint.Text,
                    Model = _model.Text,
                    ApiKey = _apiKey.Text
                };
                _store.Save(settings);
                SavedSettings = settings;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                _error.Text = exception.Message;
            }
        }
    }
}
