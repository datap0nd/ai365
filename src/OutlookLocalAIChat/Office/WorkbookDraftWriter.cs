using System;
using System.Collections;
using System.Collections.Generic;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Office
{
    // The single Excel write surface of the suite. All writes land
    // on the dedicated "AI365 Draft" worksheet: it is created at the
    // end of the workbook when missing, its previous draft content
    // is replaced, no other sheet is ever touched, and the workbook
    // is never saved - saving stays a human action.
    internal static class WorkbookDraftWriter
    {
        internal const string DraftSheetName = "AI365 Draft";
        internal const int MaxDraftRows = 200;
        internal const int MaxDraftColumns = 30;
        internal const int MaxCellCharacters = 500;

        internal static string WriteDraftSheet(
            object excelApplication,
            string title,
            IReadOnlyList<IReadOnlyList<string>> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                throw new InvalidOperationException(
                    "The draft table needs at least one row.");
            }

            dynamic application = excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                workbook = application.Workbooks.Add();
            }

            dynamic sheet = null;
            foreach (dynamic candidate in workbook.Worksheets)
            {
                if (string.Equals(
                    Convert.ToString(candidate.Name),
                    DraftSheetName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet == null)
            {
                dynamic sheets = workbook.Worksheets;
                sheet = sheets.Add(
                    Type.Missing,
                    sheets[sheets.Count]);
                sheet.Name = DraftSheetName;
            }
            else
            {
                sheet.UsedRange.ClearContents();
            }

            var boundedTitle = TextBoundary.SingleLine(
                title,
                180);
            var startRow = 1;
            if (boundedTitle.Length > 0)
            {
                sheet.Cells[1, 1].Value2 = boundedTitle;
                startRow = 3;
            }

            var rowCount = Math.Min(rows.Count, MaxDraftRows);
            var columnCount = 1;
            for (var index = 0; index < rowCount; index++)
            {
                if (rows[index] != null &&
                    rows[index].Count > columnCount)
                {
                    columnCount = Math.Min(
                        rows[index].Count,
                        MaxDraftColumns);
                }
            }

            var grid = new object[rowCount, columnCount];
            for (var row = 0; row < rowCount; row++)
            {
                var source = rows[row];
                for (var column = 0;
                     column < columnCount;
                     column++)
                {
                    grid[row, column] =
                        source != null && column < source.Count
                            ? TextBoundary.SingleLine(
                                source[column],
                                MaxCellCharacters)
                            : string.Empty;
                }
            }

            dynamic target = sheet.Range(
                sheet.Cells[startRow, 1],
                sheet.Cells[
                    startRow + rowCount - 1,
                    columnCount]);
            target.Value2 = grid;
            try
            {
                sheet.Activate();
            }
            catch
            {
            }

            return "Wrote " + rowCount + " rows x " +
                columnCount + " columns to the '" +
                DraftSheetName +
                "' sheet. Nothing was saved.";
        }

        // Converts the model-supplied JSON rows value into bounded
        // string rows, rejecting anything but arrays of arrays.
        internal static IReadOnlyList<IReadOnlyList<string>> ParseRows(
            object value)
        {
            var outer = AsEnumerable(value);
            if (outer == null)
            {
                throw new InvalidOperationException(
                    "rows must be an array of arrays of strings.");
            }

            var rows = new List<IReadOnlyList<string>>();
            foreach (var rowValue in outer)
            {
                if (rows.Count == MaxDraftRows)
                {
                    break;
                }

                var inner = AsEnumerable(rowValue);
                if (inner == null)
                {
                    throw new InvalidOperationException(
                        "rows must be an array of arrays of strings.");
                }

                var cells = new List<string>();
                foreach (var cell in inner)
                {
                    if (cells.Count == MaxDraftColumns)
                    {
                        break;
                    }

                    cells.Add(TextBoundary.SingleLine(
                        Convert.ToString(cell),
                        MaxCellCharacters));
                }

                rows.Add(cells);
            }

            return rows;
        }

        private static IEnumerable AsEnumerable(object value)
        {
            if (value == null || value is string)
            {
                return null;
            }

            return value as IEnumerable;
        }
    }
}
