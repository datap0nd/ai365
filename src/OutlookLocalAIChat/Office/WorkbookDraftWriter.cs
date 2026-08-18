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
    // is never saved - saving stays a human action. Cells starting
    // with '=' become live formulas only when DraftFormulaPolicy
    // allows them (no network, native-code, or external-workbook
    // functions); everything else lands as text.
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

            // Formula cells stay out of the bulk write: they are
            // set one by one below so a rejected or broken formula
            // degrades to text without failing the whole draft.
            var formulas =
                new List<KeyValuePair<int[], string>>();
            var grid = new object[rowCount, columnCount];
            for (var row = 0; row < rowCount; row++)
            {
                var source = rows[row];
                for (var column = 0;
                     column < columnCount;
                     column++)
                {
                    var cell =
                        source != null && column < source.Count
                            ? TextBoundary.SingleLine(
                                source[column],
                                MaxCellCharacters)
                            : string.Empty;
                    if (cell.Length > 0 && cell[0] == '=')
                    {
                        if (DraftFormulaPolicy.IsAllowedFormula(
                            cell))
                        {
                            grid[row, column] = string.Empty;
                            formulas.Add(
                                new KeyValuePair<int[], string>(
                                    new[] { row, column },
                                    cell));
                            continue;
                        }

                        // The apostrophe keeps blocked formula
                        // text visible as plain text.
                        cell = "'" + cell;
                    }

                    grid[row, column] = cell;
                }
            }

            dynamic target = sheet.Range(
                sheet.Cells[startRow, 1],
                sheet.Cells[
                    startRow + rowCount - 1,
                    columnCount]);
            target.Value2 = grid;
            var formulaCount = 0;
            foreach (var formula in formulas)
            {
                try
                {
                    sheet.Cells[
                        startRow + formula.Key[0],
                        formula.Key[1] + 1].Formula =
                        formula.Value;
                    formulaCount++;
                }
                catch
                {
                    try
                    {
                        sheet.Cells[
                            startRow + formula.Key[0],
                            formula.Key[1] + 1].Value2 =
                            "'" + formula.Value;
                    }
                    catch
                    {
                    }
                }
            }

            ApplyDraftFormatting(
                sheet,
                boundedTitle,
                startRow,
                rowCount,
                target);
            try
            {
                sheet.Activate();
            }
            catch
            {
            }

            return "Wrote " + rowCount + " rows x " +
                columnCount + " columns" +
                (formulaCount > 0
                    ? " including " + formulaCount +
                      " live formulas"
                    : string.Empty) +
                " to the '" +
                DraftSheetName +
                "' sheet. Nothing was saved.";
        }

        // Cosmetic polish for the draft sheet: bold title, bold
        // header row with a divider, and autofitted columns. Any
        // failure here must never fail the draft itself.
        private static void ApplyDraftFormatting(
            dynamic sheet,
            string boundedTitle,
            int startRow,
            int rowCount,
            dynamic target)
        {
            try
            {
                if (boundedTitle.Length > 0)
                {
                    dynamic titleCell = sheet.Cells[1, 1];
                    titleCell.Font.Bold = true;
                    titleCell.Font.Size = 12;
                }

                if (rowCount > 1)
                {
                    dynamic header = target.Rows[1];
                    header.Font.Bold = true;
                    // 9 = xlEdgeBottom, 1 = xlContinuous.
                    header.Borders[9].LineStyle = 1;
                }

                target.EntireColumn.AutoFit();
            }
            catch
            {
            }
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
