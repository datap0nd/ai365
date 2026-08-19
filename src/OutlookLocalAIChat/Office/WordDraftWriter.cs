using System;
using System.Collections.Generic;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Office
{
    // The single Word write surface of the suite. Every draft goes
    // into a brand-new, unsaved document opened in its own window
    // with the [AI365 draft] marker as its heading; the user's
    // existing documents are never touched and nothing is ever
    // saved - saving stays a human action. Draft text supports a
    // small layout vocabulary (#/##/### headings, -/1. lists,
    // **bold**, and | cell | cell | rows that become real Word
    // tables); styling is cosmetic and can never fail the draft.
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

            var blocks = DraftTextLayout.ParseBlocks(boundedBody);
            dynamic application = wordApplication;
            dynamic document = application.Documents.Add();
            AppendParagraph(
                document,
                new DraftTextLayout.Paragraph(
                    heading,
                    0,
                    null),
                StyleHeading1);
            var tables = 0;
            foreach (var block in blocks)
            {
                var table = block as DraftTextLayout.Table;
                if (table != null)
                {
                    if (AppendTable(document, table))
                    {
                        tables++;
                    }
                    else
                    {
                        // A failed table insert degrades to plain
                        // rows so no content is ever lost.
                        foreach (var row in table.Rows)
                        {
                            AppendParagraph(
                                document,
                                new DraftTextLayout.Paragraph(
                                    string.Join(
                                        "  |  ",
                                        row),
                                    0,
                                    null),
                                0);
                        }
                    }

                    continue;
                }

                var paragraph =
                    (DraftTextLayout.Paragraph)block;
                AppendParagraph(
                    document,
                    paragraph,
                    StyleFor(paragraph.Kind));
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
                boundedBody.Length + " characters" +
                (tables > 0
                    ? " with " + tables +
                      (tables == 1 ? " table" : " tables")
                    : string.Empty) +
                ". Nothing was saved.";
        }

        // Inserts one paragraph at the end of the document; style
        // and bold runs are cosmetic and can never fail the draft.
        private static void AppendParagraph(
            dynamic document,
            DraftTextLayout.Paragraph paragraph,
            int style)
        {
            var end = (int)document.Content.End - 1;
            dynamic range = document.Range(end, end);
            range.Text = paragraph.Text + "\r";
            if (style != 0)
            {
                try
                {
                    range.Style = style;
                }
                catch
                {
                }
            }

            foreach (var bold in paragraph.BoldRanges)
            {
                try
                {
                    document.Range(
                        end + bold.Start,
                        end + bold.Start + bold.Length)
                        .Font.Bold = 1;
                }
                catch
                {
                }
            }
        }

        // Inserts one real Word table at the end of the document.
        // Returns false when the table could not be created so the
        // caller can fall back to plain text rows.
        private static bool AppendTable(
            dynamic document,
            DraftTextLayout.Table table)
        {
            var rowCount = table.Rows.Count;
            if (rowCount == 0)
            {
                return true;
            }

            var columnCount = 1;
            foreach (var row in table.Rows)
            {
                if (row.Count > columnCount)
                {
                    columnCount = row.Count;
                }
            }

            try
            {
                var end = (int)document.Content.End - 1;
                dynamic anchor = document.Range(end, end);
                dynamic wordTable = document.Tables.Add(
                    anchor,
                    rowCount,
                    columnCount);
                for (var row = 0; row < rowCount; row++)
                {
                    var cells = table.Rows[row];
                    for (var column = 0;
                         column < cells.Count;
                         column++)
                    {
                        wordTable.Cell(row + 1, column + 1)
                            .Range.Text = TextBoundary.SingleLine(
                                cells[column],
                                500);
                    }
                }

                try
                {
                    // 1 = wdAutoFitContent.
                    wordTable.AutoFitBehavior(1);
                    wordTable.Rows[1].Range.Font.Bold = 1;
                    // 1 = enable default single-line borders.
                    wordTable.Borders.Enable = 1;
                }
                catch
                {
                }

                try
                {
                    wordTable.Style = "Grid Table 4 - Accent 1";
                }
                catch
                {
                    // The built-in style name is language-specific;
                    // the manual borders above already keep the
                    // table readable.
                }

                // A spacer paragraph after the table keeps the next
                // block out of the table.
                var after = (int)document.Content.End - 1;
                document.Range(after, after).Text = "\r";
                return true;
            }
            catch
            {
                return false;
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
