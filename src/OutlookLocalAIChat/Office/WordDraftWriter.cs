using System;
using System.Collections.Generic;
using System.Text;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Office
{
    // The single Word write surface of the suite. Every draft goes
    // into a brand-new, unsaved document opened in its own window
    // with the [AI365 draft] marker as its heading; the user's
    // existing documents are never touched and nothing is ever
    // saved - saving stays a human action. Draft text supports a
    // small layout vocabulary (#/##/### headings, -/1. lists, and
    // **bold**) rendered as real Word styles; styling is cosmetic
    // and can never fail the draft.
    internal static class WordDraftWriter
    {
        internal const string DraftMarker = "[AI365 draft]";
        internal const int MaxDraftCharacters = 48000;
        internal const int MaxTitleCharacters = 180;

        // WdBuiltinStyle ids work in every localized Word.
        private const int StyleHeading1 = -2;
        private const int StyleHeading2 = -3;
        private const int StyleHeading3 = -4;
        private const int StyleListBullet = -49;
        private const int StyleListNumber = -50;

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

            var paragraphs = DraftTextLayout.Parse(boundedBody);
            var text = new StringBuilder(
                boundedBody.Length + heading.Length + 16);
            text.Append(heading);
            foreach (var paragraph in paragraphs)
            {
                // Word paragraph marks are carriage returns.
                text.Append('\r');
                text.Append(paragraph.Text);
            }

            dynamic application = wordApplication;
            dynamic document = application.Documents.Add();
            document.Content.Text = text.ToString();
            ApplyStyles(document, heading, paragraphs);
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

        private static void ApplyStyles(
            dynamic document,
            string heading,
            IReadOnlyList<DraftTextLayout.Paragraph> paragraphs)
        {
            try
            {
                document.Paragraphs[1].Range.Style = StyleHeading1;
            }
            catch
            {
                // Styling is cosmetic and must never fail the
                // draft.
            }

            var offset = heading.Length + 1;
            for (var index = 0; index < paragraphs.Count; index++)
            {
                var paragraph = paragraphs[index];
                var style = StyleFor(paragraph.Kind);
                if (style != 0)
                {
                    try
                    {
                        document.Paragraphs[index + 2]
                            .Range.Style = style;
                    }
                    catch
                    {
                    }
                }

                foreach (var range in paragraph.BoldRanges)
                {
                    try
                    {
                        document.Range(
                            offset + range.Start,
                            offset + range.Start + range.Length)
                            .Font.Bold = 1;
                    }
                    catch
                    {
                    }
                }

                offset += paragraph.Text.Length + 1;
            }
        }

        private static int StyleFor(int kind)
        {
            switch (kind)
            {
                case DraftTextLayout.KindHeading1:
                    return StyleHeading2;
                case DraftTextLayout.KindHeading2:
                    return StyleHeading3;
                case DraftTextLayout.KindHeading3:
                    return StyleHeading3;
                case DraftTextLayout.KindBullet:
                    return StyleListBullet;
                case DraftTextLayout.KindNumbered:
                    return StyleListNumber;
                default:
                    return 0;
            }
        }
    }
}
