using System;
using System.Collections.Generic;
using OutlookLocalAIChat.Office;

namespace OutlookLocalAIChat.Chat
{
    // Cross-application draft tools for the Excel and PowerPoint
    // panes: hand content to a sibling Office app as a clearly
    // marked, unsent, unsaved draft. Email drafts open for review in
    // Outlook and can never be sent; slide and sheet drafts follow
    // the same marked-draft rules as the in-host writers. All three
    // are offered only when the user's own prompt authorized a
    // draft, and they share the same one-shot permission.
    public static class CrossAppToolCatalog
    {
        public const string CreateEmailDraft = "create_email_draft";
        public const string SendToPowerPoint = "send_to_powerpoint";
        public const string SendToExcel = "send_to_excel";

        public static List<ChatToolDefinition> CreateDefinitions(
            string hostKind)
        {
            var definitions = new List<ChatToolDefinition>
            {
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = CreateEmailDraft,
                        description =
                            "Open one unsent Outlook email draft with the " +
                            "given recipients, subject, and body for the " +
                            "user to review. The draft is never sent and " +
                            "sending is impossible. When the current file " +
                            "is saved on disk it can be attached with " +
                            "attach_current_file. For a visual email use " +
                            "only these layout lines in body: # heading, " +
                            "## subheading, - list item, 1. numbered item, " +
                            "and --- divider.",
                        parameters = ToolSchema.Build(
                            new Dictionary<string, object>
                            {
                                {
                                    "to",
                                    ToolSchema.String(
                                        "Recipient addresses separated by " +
                                        "semicolons. May be empty for the " +
                                        "user to fill in.")
                                },
                                {
                                    "cc",
                                    ToolSchema.String(
                                        "Cc addresses separated by semicolons.")
                                },
                                {
                                    "subject",
                                    ToolSchema.String(
                                        "Email subject line.")
                                },
                                {
                                    "body",
                                    ToolSchema.String(
                                        "Plain-text email body using only the " +
                                        "allowed layout lines.")
                                },
                                {
                                    "attach_current_file",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "boolean" },
                                        {
                                            "description",
                                            "Attach the current document file " +
                                            "when it is saved on disk."
                                        }
                                    }
                                }
                            },
                            "body")
                    }
                }
            };

            if (string.Equals(
                hostKind,
                "excel",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SendToPowerPoint,
                        description =
                            "Append clearly marked [AI365 draft] slides to " +
                            "the active PowerPoint presentation (a new " +
                            "unsaved presentation is opened when none is " +
                            "active). Existing slides are never modified " +
                            "and nothing is saved. At most " +
                            PresentationDraftWriter.MaxDraftSlides +
                            " slides per call.",
                        parameters =
                            PresentationToolCatalog.DraftDefinition()
                                .function.parameters as
                                Dictionary<string, object>
                    }
                });
            }

            if (string.Equals(
                hostKind,
                "powerpoint",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SendToExcel,
                        description =
                            "Write a table into the dedicated 'AI365 Draft' " +
                            "worksheet of the active Excel workbook (a new " +
                            "unsaved workbook is opened when none is " +
                            "active). No other sheet is ever touched and " +
                            "nothing is saved.",
                        parameters =
                            WorkbookToolCatalog.DraftDefinition()
                                .function.parameters as
                                Dictionary<string, object>
                    }
                });
            }

            return definitions;
        }

        public static bool IsCrossAppTool(string name)
        {
            return string.Equals(
                       name,
                       CreateEmailDraft,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       name,
                       SendToPowerPoint,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       name,
                       SendToExcel,
                       StringComparison.Ordinal);
        }
    }
}
