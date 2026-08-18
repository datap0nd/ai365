using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.Office
{
    // Every mutating tool of the Excel and PowerPoint panes runs
    // through this host, behind the same one-shot authorization the
    // Outlook draft host uses: the user's own prompt must have
    // authorized a draft, the permission is consumed on the first
    // attempt, and the tool must be the only call in its response.
    // The write surfaces are all clearly marked drafts - the
    // 'AI365 Draft' worksheet, appended '[AI365 draft]' slides, and
    // unsent Outlook email drafts. Nothing is ever saved or sent.
    public sealed class DocumentDraftHost : IDisposable
    {
        private static readonly HashSet<string> EmailArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "to",
                "cc",
                "subject",
                "body",
                "attach_current_file"
            };

        private static readonly HashSet<string> SheetArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "title",
                "rows"
            };

        private static readonly HashSet<string> SlideArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "slides"
            };

        private static readonly HashSet<string> DocumentArguments =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "title",
                "body"
            };

        private readonly string _hostKind;
        private readonly object _hostApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private DraftSession _emailDraft;
        private string _latestUserPrompt = string.Empty;

        public DocumentDraftHost(
            string hostKind,
            object hostApplication)
        {
            _hostKind = hostKind ?? string.Empty;
            _hostApplication = hostApplication ??
                throw new ArgumentNullException(
                    nameof(hostApplication));
        }

        public static bool IsDraftTool(
            string hostKind,
            string name)
        {
            if (CrossAppToolCatalog.IsCrossAppTool(name))
            {
                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToPowerPoint,
                        StringComparison.Ordinal))
                {
                    return hostKind == "excel" ||
                           hostKind == "word";
                }

                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToExcel,
                        StringComparison.Ordinal))
                {
                    return hostKind == "powerpoint" ||
                           hostKind == "word";
                }

                if (string.Equals(
                        name,
                        CrossAppToolCatalog.SendToWord,
                        StringComparison.Ordinal))
                {
                    return hostKind == "excel" ||
                           hostKind == "powerpoint";
                }

                return true;
            }

            if (hostKind == "excel")
            {
                return WorkbookToolCatalog.IsDraftTool(name);
            }

            if (hostKind == "powerpoint")
            {
                return PresentationToolCatalog.IsDraftTool(name);
            }

            if (hostKind == "word")
            {
                return WordToolCatalog.IsDraftTool(name);
            }

            return false;
        }

        public MailboxToolResult Execute(
            ChatToolCall call,
            OneShotDraftAuthorization authorization,
            bool isOnlyToolCall,
            string userPrompt = null)
        {
            _latestUserPrompt = userPrompt ?? string.Empty;
            var name = call?.function?.name;
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !IsDraftTool(_hostKind, name))
            {
                return Error(
                    call?.id,
                    authorization,
                    "DRAFT_TOOL_CALL_INVALID",
                    "The model returned an invalid draft tool call.");
            }

            if (!isOnlyToolCall)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
                    name + " must be the only tool call in its response.");
            }

            IDictionary<string, object> arguments;
            try
            {
                arguments = ToolArguments.Parse(
                    _serializer,
                    call.function.arguments);
                RequireAllowedArguments(arguments, name);
            }
            catch (Exception exception)
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_ARGUMENTS_INVALID",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_ARGUMENTS_INVALID"));
            }

            if (authorization == null ||
                !authorization.CanCreate ||
                !authorization.TryConsume())
            {
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "No unused local draft permission is available for this request.");
            }

            try
            {
                string status;
                if (string.Equals(
                    name,
                    CrossAppToolCatalog.CreateEmailDraft,
                    StringComparison.Ordinal))
                {
                    status = CreateEmailDraft(arguments);
                }
                else if (string.Equals(
                             name,
                             WorkbookToolCatalog.WriteDraftSheet,
                             StringComparison.Ordinal))
                {
                    status = WorkbookDraftWriter.WriteDraftSheet(
                        _hostApplication,
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        ParsedRows(arguments));
                }
                else if (string.Equals(
                             name,
                             PresentationToolCatalog.AddDraftSlides,
                             StringComparison.Ordinal))
                {
                    status = PresentationDraftWriter.AddDraftSlides(
                        _hostApplication,
                        ParsedSlides(arguments));
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToPowerPoint,
                             StringComparison.Ordinal))
                {
                    status = PresentationDraftWriter.AddDraftSlides(
                        GetSiblingApplication(
                            "PowerPoint.Application"),
                        ParsedSlides(arguments));
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToExcel,
                             StringComparison.Ordinal))
                {
                    status = WorkbookDraftWriter.WriteDraftSheet(
                        GetSiblingApplication(
                            "Excel.Application"),
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        ParsedRows(arguments));
                }
                else if (string.Equals(
                             name,
                             WordToolCatalog.WriteDraftDocument,
                             StringComparison.Ordinal))
                {
                    status = WordDraftWriter.WriteDraftDocument(
                        _hostApplication,
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        GetLongString(arguments, "body"));
                }
                else if (string.Equals(
                             name,
                             CrossAppToolCatalog.SendToWord,
                             StringComparison.Ordinal))
                {
                    status = WordDraftWriter.WriteDraftDocument(
                        GetSiblingApplication(
                            "Word.Application"),
                        ToolArguments.GetString(
                            arguments,
                            "title",
                            string.Empty),
                        GetLongString(arguments, "body"));
                }
                else
                {
                    throw new InvalidOperationException(
                        "The requested draft tool is not allowed.");
                }

                authorization.MarkCreated();
                return new MailboxToolResult(
                    call.id,
                    _serializer.Serialize(
                        new Dictionary<string, object>
                        {
                            { "ok", true },
                            { "action", name },
                            { "saved", false },
                            { "sent", false },
                            {
                                "status",
                                TextBoundary.PlainText(status, 600)
                            }
                        }),
                    TextBoundary.SingleLine(status, 300));
            }
            catch (Exception exception)
            {
                Log.Error("DocumentDraft." + name, exception);
                return Error(
                    call.id,
                    authorization,
                    "DRAFT_CREATION_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "DRAFT_CREATION_FAILED"));
            }
        }

        public void Dispose()
        {
            _emailDraft?.Dispose();
            _emailDraft = null;
        }

        // Long text arguments (draft bodies) must not ride through
        // the 1000-character generic string reader.
        private static string GetLongString(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.PlainText(
                    Convert.ToString(value),
                    WordDraftWriter.MaxDraftCharacters)
                : string.Empty;
        }

        private static IReadOnlyList<IReadOnlyList<string>>
            ParsedRows(IDictionary<string, object> arguments)
        {
            object rowsValue;
            arguments.TryGetValue("rows", out rowsValue);
            return WorkbookDraftWriter.ParseRows(rowsValue);
        }

        private static
            IReadOnlyList<PresentationDraftWriter.DraftSlide>
            ParsedSlides(IDictionary<string, object> arguments)
        {
            object slidesValue;
            arguments.TryGetValue("slides", out slidesValue);
            return PresentationDraftWriter.ParseSlides(slidesValue);
        }

        // Opens one unsent Outlook draft for review via the same
        // DraftService the Outlook pane uses: the draft is saved to
        // Drafts and displayed, never sent - sending stays a human
        // action in Outlook.
        private string CreateEmailDraft(
            IDictionary<string, object> arguments)
        {
            var body = TextBoundary.PlainText(
                GetLongString(arguments, "body"),
                TextBoundary.MaxAssistantCharacters);
            if (body.Length == 0)
            {
                throw new InvalidOperationException(
                    "A non-empty draft body is required.");
            }

            var outlook = GetSiblingApplication(
                "Outlook.Application");
            var drafts = new DraftService(outlook);
            var to = ToolArguments.GetString(
                arguments,
                "to",
                string.Empty);
            var cc = ToolArguments.GetString(
                arguments,
                "cc",
                string.Empty);
            _emailDraft?.Dispose();
            _emailDraft = drafts.CreateNewDraft(
                body,
                new string[0],
                ToolArguments.GetString(
                    arguments,
                    "subject",
                    string.Empty),
                to,
                cc);
            var attached = false;
            if (ToolArguments.GetBoolean(
                arguments,
                "attach_current_file"))
            {
                var path = CurrentDocumentPath();
                if (path.Length > 0)
                {
                    _emailDraft.AttachFile(path);
                    attached = true;
                }
            }

            return "One unsent Outlook draft was opened for " +
                "review" +
                (attached
                    ? " with the current file attached"
                    : string.Empty) +
                ". AI365 cannot send it." +
                RecipientIntentCheck.Warn(
                    to,
                    cc,
                    _latestUserPrompt);
        }

        // Full path of the host document when it exists on disk;
        // empty for unsaved documents.
        private string CurrentDocumentPath()
        {
            try
            {
                dynamic application = _hostApplication;
                dynamic document;
                if (_hostKind == "excel")
                {
                    document = application.ActiveWorkbook;
                }
                else if (_hostKind == "word")
                {
                    document = application.ActiveDocument;
                }
                else
                {
                    document = application.ActivePresentation;
                }
                if (document == null)
                {
                    return string.Empty;
                }

                var folder = Convert.ToString(document.Path) ??
                    string.Empty;
                if (folder.Length == 0)
                {
                    return string.Empty;
                }

                return Convert.ToString(document.FullName) ??
                    string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        // Reuses the running sibling Office application when there
        // is one; otherwise starts it visibly so the draft opens in
        // front of the user.
        private static object GetSiblingApplication(string progId)
        {
            try
            {
                return Marshal.GetActiveObject(progId);
            }
            catch
            {
            }

            var type = Type.GetTypeFromProgID(progId);
            if (type == null)
            {
                throw new InvalidOperationException(
                    progId.Split('.')[0] +
                    " is not installed on this computer.");
            }

            var application = Activator.CreateInstance(type);
            try
            {
                dynamic visible = application;
                visible.Visible = true;
            }
            catch
            {
                // Outlook has no Visible property; drafts are
                // displayed by the draft session itself.
            }

            return application;
        }

        private void RequireAllowedArguments(
            IDictionary<string, object> arguments,
            string name)
        {
            ISet<string> allowed;
            if (string.Equals(
                name,
                CrossAppToolCatalog.CreateEmailDraft,
                StringComparison.Ordinal))
            {
                allowed = EmailArguments;
            }
            else if (
                string.Equals(
                    name,
                    WorkbookToolCatalog.WriteDraftSheet,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    CrossAppToolCatalog.SendToExcel,
                    StringComparison.Ordinal))
            {
                allowed = SheetArguments;
            }
            else if (
                string.Equals(
                    name,
                    WordToolCatalog.WriteDraftDocument,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    CrossAppToolCatalog.SendToWord,
                    StringComparison.Ordinal))
            {
                allowed = DocumentArguments;
            }
            else
            {
                allowed = SlideArguments;
            }

            var unexpected = arguments.Keys
                .FirstOrDefault(key => !allowed.Contains(key));
            if (unexpected != null)
            {
                throw new InvalidOperationException(
                    "Unexpected draft argument: " + unexpected);
            }
        }

        private MailboxToolResult Error(
            string callId,
            OneShotDraftAuthorization authorization,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        },
                        {
                            "permission_consumed",
                            authorization != null &&
                            authorization.IsConsumed
                        }
                    }),
                "[" + code + "] " +
                TextBoundary.PlainText(message, 600));
        }
    }
}
