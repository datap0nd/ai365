using System;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Office
{
    // The single Word write surface of the suite. Every draft goes
    // into a brand-new, unsaved document opened in its own window
    // with the [AI365 draft] marker as its heading; the user's
    // existing documents are never touched and nothing is ever
    // saved - saving stays a human action.
    internal static class WordDraftWriter
    {
        internal const string DraftMarker = "[AI365 draft]";
        internal const int MaxDraftCharacters = 48000;
        internal const int MaxTitleCharacters = 180;

        internal static string WriteDraftDocument(
            object wordApplication,
            string title,
            string body)
        {
            var boundedBody = TextBoundary.PlainText(
                body,
                MaxDraftCharacters);
            if (boundedBody.Length == 0)
            {
                throw new InvalidOperationException(
                    "A non-empty draft body is required.");
            }

            var heading = DraftMarker;
            var boundedTitle = TextBoundary.SingleLine(
                title,
                MaxTitleCharacters);
            if (boundedTitle.Length > 0)
            {
                heading += " " + boundedTitle;
            }

            dynamic application = wordApplication;
            dynamic document = application.Documents.Add();
            // Word paragraph marks are carriage returns.
            document.Content.Text =
                heading + "\r" +
                boundedBody.Replace("\r\n", "\n")
                    .Replace('\n', '\r');
            try
            {
                document.Paragraphs[1].Range.Style = "Heading 1";
            }
            catch
            {
                // Styling is cosmetic; a missing style name in a
                // localized Word must never fail the draft.
            }

            try
            {
                document.Activate();
                application.Visible = true;
            }
            catch
            {
            }

            return "Opened a new unsaved Word draft document of " +
                boundedBody.Length + " characters. Nothing was " +
                "saved.";
        }
    }
}
