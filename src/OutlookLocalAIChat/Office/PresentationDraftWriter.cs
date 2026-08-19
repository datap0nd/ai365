using System;
using System.Collections;
using System.Collections.Generic;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Office
{
    // The single PowerPoint write surface of the suite. Draft slides
    // are appended at the end of the presentation, every added title
    // carries the [AI365 draft] marker, existing slides are never
    // modified, and the file is never saved - saving stays a human
    // action.
    internal static class PresentationDraftWriter
    {
        internal const string DraftMarker = "[AI365 draft]";
        internal const int MaxDraftSlides = 10;
        internal const int MaxBulletsPerSlide = 12;
        internal const int MaxTitleCharacters = 200;
        internal const int MaxBulletCharacters = 300;

        internal const int MaxChartCategories = 24;
        internal const int MaxChartSeries = 5;

        internal sealed class DraftSlide
        {
            internal DraftSlide(
                string title,
                IReadOnlyList<DraftBullet> bullets,
                DraftChart chart)
            {
                Title = title ?? string.Empty;
                Bullets = bullets ?? new DraftBullet[0];
                Chart = chart;
            }

            internal string Title { get; }

            internal IReadOnlyList<DraftBullet> Bullets { get; }

            internal DraftChart Chart { get; }
        }

        // A native chart drawn on the slide from bounded data the
        // model supplies: categories down the side, one to five
        // named series of numbers.
        internal sealed class DraftChart
        {
            internal DraftChart(
                int typeCode,
                string title,
                IReadOnlyList<string> categories,
                IReadOnlyList<DraftChartSeries> series)
            {
                TypeCode = typeCode;
                Title = title ?? string.Empty;
                Categories = categories ?? new string[0];
                Series = series ?? new DraftChartSeries[0];
            }

            internal int TypeCode { get; }

            internal string Title { get; }

            internal IReadOnlyList<string> Categories { get; }

            internal IReadOnlyList<DraftChartSeries> Series
            {
                get;
            }
        }

        internal sealed class DraftChartSeries
        {
            internal DraftChartSeries(
                string name,
                IReadOnlyList<double> values)
            {
                Name = name ?? string.Empty;
                Values = values ?? new double[0];
            }

            internal string Name { get; }

            internal IReadOnlyList<double> Values { get; }
        }

        // A bullet line with its outline level (1-5). Sub-bullets
        // are written with two leading spaces per extra level.
        internal sealed class DraftBullet
        {
            internal DraftBullet(string text, int level)
            {
                Text = text ?? string.Empty;
                Level = level < 1 ? 1 : (level > 5 ? 5 : level);
            }

            internal string Text { get; }

            internal int Level { get; }
        }

        internal static string AddDraftSlides(
            object powerPointApplication,
            IReadOnlyList<DraftSlide> slides)
        {
            if (slides == null || slides.Count == 0)
            {
                throw new InvalidOperationException(
                    "At least one draft slide is required.");
            }

            dynamic application = powerPointApplication;
            dynamic presentation = null;
            try
            {
                presentation = application.ActivePresentation;
            }
            catch
            {
            }

            if (presentation == null)
            {
                // msoTrue window so the new unsaved deck is visible.
                presentation = application.Presentations.Add(-1);
            }

            var added = 0;
            var charts = 0;
            foreach (var slide in slides)
            {
                if (added == MaxDraftSlides)
                {
                    break;
                }

                var index = (int)presentation.Slides.Count + 1;
                // 2 = ppLayoutText (title plus body placeholder);
                // 11 = ppLayoutTitleOnly for chart-only slides so
                // the chart gets the full body area.
                dynamic created = presentation.Slides.Add(
                    index,
                    slide.Chart != null &&
                    slide.Bullets.Count == 0
                        ? 11
                        : 2);
                var titleText = DraftMarker + " " +
                    TextBoundary.SingleLine(
                        slide.Title,
                        MaxTitleCharacters);
                try
                {
                    created.Shapes.Title.TextFrame.TextRange.Text =
                        titleText;
                }
                catch
                {
                }

                if (slide.Bullets.Count > 0)
                {
                    var lines = new List<DraftBullet>();
                    foreach (var bullet in slide.Bullets)
                    {
                        if (lines.Count == MaxBulletsPerSlide)
                        {
                            break;
                        }

                        if (bullet.Text.Length > 0)
                        {
                            lines.Add(bullet);
                        }
                    }

                    if (lines.Count > 0)
                    {
                        try
                        {
                            var texts = new List<string>();
                            foreach (var line in lines)
                            {
                                texts.Add(line.Text);
                            }

                            dynamic textRange = created.Shapes
                                .Placeholders[2]
                                .TextFrame.TextRange;
                            textRange.Text =
                                string.Join("\r", texts);
                            for (var line = 0;
                                 line < lines.Count;
                                 line++)
                            {
                                if (lines[line].Level > 1)
                                {
                                    try
                                    {
                                        textRange
                                            .Paragraphs(line + 1)
                                            .IndentLevel =
                                            lines[line].Level;
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                }

                if (slide.Chart != null)
                {
                    if (AddChartToSlide(
                        created,
                        presentation,
                        slide.Chart,
                        slide.Bullets.Count > 0))
                    {
                        charts++;
                    }
                }

                added++;
            }

            return "Appended " + added +
                " marked draft slides" +
                (charts > 0
                    ? " with " + charts +
                      (charts == 1
                          ? " native chart"
                          : " native charts")
                    : string.Empty) +
                " at the end of the presentation. Nothing was " +
                "saved.";
        }

        // Draws one native chart on the slide and fills its data
        // grid from the bounded model-supplied numbers. The data
        // grid workbook is the chart's own embedded store inside
        // the unsaved draft presentation - closing it only closes
        // the editing grid; no user file is touched or saved.
        private static bool AddChartToSlide(
            dynamic slide,
            dynamic presentation,
            DraftChart chart,
            bool hasBullets)
        {
            try
            {
                double slideWidth =
                    presentation.PageSetup.SlideWidth;
                double slideHeight =
                    presentation.PageSetup.SlideHeight;
                var left = hasBullets
                    ? slideWidth * 0.52
                    : slideWidth * 0.08;
                var width = hasBullets
                    ? slideWidth * 0.42
                    : slideWidth * 0.84;
                var top = slideHeight * 0.24;
                var height = slideHeight * 0.62;
                dynamic shape = slide.Shapes.AddChart2(
                    -1,
                    chart.TypeCode,
                    (float)left,
                    (float)top,
                    (float)width,
                    (float)height,
                    true);
                dynamic slideChart = shape.Chart;
                slideChart.ChartData.Activate();
                dynamic dataWorkbook =
                    slideChart.ChartData.Workbook;
                dynamic dataSheet =
                    dataWorkbook.Worksheets[1];
                dataSheet.Cells[1, 1].Value2 = " ";
                for (var series = 0;
                     series < chart.Series.Count;
                     series++)
                {
                    dataSheet.Cells[1, series + 2].Value2 =
                        TextBoundary.SingleLine(
                            chart.Series[series].Name,
                            80);
                }

                for (var category = 0;
                     category < chart.Categories.Count;
                     category++)
                {
                    dataSheet.Cells[category + 2, 1].Value2 =
                        TextBoundary.SingleLine(
                            chart.Categories[category],
                            80);
                    for (var series = 0;
                         series < chart.Series.Count;
                         series++)
                    {
                        var values =
                            chart.Series[series].Values;
                        dataSheet.Cells[
                            category + 2,
                            series + 2].Value2 =
                            category < values.Count
                                ? values[category]
                                : 0d;
                    }
                }

                try
                {
                    dataSheet.ListObjects["Table1"].Resize(
                        dataSheet.Range(
                            dataSheet.Cells[1, 1],
                            dataSheet.Cells[
                                chart.Categories.Count + 1,
                                chart.Series.Count + 1]));
                }
                catch
                {
                }

                try
                {
                    if (chart.Title.Length > 0)
                    {
                        slideChart.HasTitle = true;
                        slideChart.ChartTitle.Text =
                            TextBoundary.SingleLine(
                                chart.Title,
                                180);
                    }
                }
                catch
                {
                }

                try
                {
                    dataWorkbook.Close(true);
                }
                catch
                {
                    // Leaving the data grid window open is only
                    // cosmetic; the chart itself already holds the
                    // data.
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        // Converts the model-supplied JSON slides value into bounded
        // draft slides, rejecting anything but arrays of objects.
        internal static IReadOnlyList<DraftSlide> ParseSlides(
            object value)
        {
            var outer = value as IEnumerable;
            if (outer == null || value is string)
            {
                throw new InvalidOperationException(
                    "slides must be an array of objects.");
            }

            var slides = new List<DraftSlide>();
            foreach (var slideValue in outer)
            {
                if (slides.Count == MaxDraftSlides)
                {
                    break;
                }

                var map = slideValue as
                    IDictionary<string, object>;
                if (map == null)
                {
                    throw new InvalidOperationException(
                        "Each slide must be an object with a title.");
                }

                object titleValue;
                map.TryGetValue("title", out titleValue);
                var title = TextBoundary.SingleLine(
                    Convert.ToString(titleValue),
                    MaxTitleCharacters);
                if (title.Length == 0)
                {
                    throw new InvalidOperationException(
                        "Each slide needs a non-empty title.");
                }

                var bullets = new List<DraftBullet>();
                object bulletsValue;
                if (map.TryGetValue("bullets", out bulletsValue))
                {
                    var list = bulletsValue as IEnumerable;
                    if (list != null && !(bulletsValue is string))
                    {
                        foreach (var bullet in list)
                        {
                            if (bullets.Count ==
                                MaxBulletsPerSlide)
                            {
                                break;
                            }

                            var raw = Convert.ToString(bullet) ??
                                string.Empty;
                            var leading = 0;
                            while (leading < raw.Length &&
                                   raw[leading] == ' ')
                            {
                                leading++;
                            }

                            bullets.Add(new DraftBullet(
                                TextBoundary.SingleLine(
                                    raw,
                                    MaxBulletCharacters),
                                1 + leading / 2));
                        }
                    }
                }

                slides.Add(new DraftSlide(
                    title,
                    bullets,
                    ParseChart(map)));
            }

            return slides;
        }

        // Reads the optional chart object of one slide; anything
        // malformed simply yields no chart rather than failing the
        // whole draft.
        private static DraftChart ParseChart(
            IDictionary<string, object> slideMap)
        {
            object chartValue;
            if (!slideMap.TryGetValue("chart", out chartValue))
            {
                return null;
            }

            var map = chartValue as IDictionary<string, object>;
            if (map == null)
            {
                return null;
            }

            object typeValue;
            object titleValue;
            object categoriesValue;
            object seriesValue;
            map.TryGetValue("type", out typeValue);
            map.TryGetValue("title", out titleValue);
            map.TryGetValue("categories", out categoriesValue);
            map.TryGetValue("series", out seriesValue);

            var categories = new List<string>();
            var categoryList = categoriesValue as IEnumerable;
            if (categoryList != null &&
                !(categoriesValue is string))
            {
                foreach (var category in categoryList)
                {
                    if (categories.Count == MaxChartCategories)
                    {
                        break;
                    }

                    categories.Add(TextBoundary.SingleLine(
                        Convert.ToString(category),
                        80));
                }
            }

            var series = new List<DraftChartSeries>();
            var seriesList = seriesValue as IEnumerable;
            if (seriesList != null && !(seriesValue is string))
            {
                foreach (var entry in seriesList)
                {
                    if (series.Count == MaxChartSeries)
                    {
                        break;
                    }

                    var entryMap = entry as
                        IDictionary<string, object>;
                    if (entryMap == null)
                    {
                        continue;
                    }

                    object nameValue;
                    object valuesValue;
                    entryMap.TryGetValue("name", out nameValue);
                    entryMap.TryGetValue(
                        "values",
                        out valuesValue);
                    var values = new List<double>();
                    var valueList = valuesValue as IEnumerable;
                    if (valueList != null &&
                        !(valuesValue is string))
                    {
                        foreach (var value in valueList)
                        {
                            if (values.Count ==
                                MaxChartCategories)
                            {
                                break;
                            }

                            double parsed;
                            values.Add(double.TryParse(
                                Convert.ToString(
                                    value,
                                    System.Globalization
                                        .CultureInfo
                                        .InvariantCulture),
                                System.Globalization
                                    .NumberStyles.Any,
                                System.Globalization
                                    .CultureInfo
                                    .InvariantCulture,
                                out parsed)
                                ? parsed
                                : 0d);
                        }
                    }

                    series.Add(new DraftChartSeries(
                        TextBoundary.SingleLine(
                            Convert.ToString(nameValue),
                            80),
                        values));
                }
            }

            if (categories.Count == 0 || series.Count == 0)
            {
                return null;
            }

            return new DraftChart(
                DraftChartTypes.Resolve(
                    Convert.ToString(typeValue)),
                TextBoundary.SingleLine(
                    Convert.ToString(titleValue),
                    180),
                categories,
                series);
        }
    }
}
