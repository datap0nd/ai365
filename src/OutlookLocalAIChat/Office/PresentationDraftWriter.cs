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

        internal sealed class DraftSlide
        {
            internal DraftSlide(
                string title,
                IReadOnlyList<string> bullets)
            {
                Title = title ?? string.Empty;
                Bullets = bullets ?? new string[0];
            }

            internal string Title { get; }

            internal IReadOnlyList<string> Bullets { get; }
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
            foreach (var slide in slides)
            {
                if (added == MaxDraftSlides)
                {
                    break;
                }

                var index = (int)presentation.Slides.Count + 1;
                // 2 = ppLayoutText: title plus body placeholder.
                dynamic created = presentation.Slides.Add(
                    index,
                    2);
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
                    var lines = new List<string>();
                    foreach (var bullet in slide.Bullets)
                    {
                        if (lines.Count == MaxBulletsPerSlide)
                        {
                            break;
                        }

                        var bounded = TextBoundary.SingleLine(
                            bullet,
                            MaxBulletCharacters);
                        if (bounded.Length > 0)
                        {
                            lines.Add(bounded);
                        }
                    }

                    if (lines.Count > 0)
                    {
                        try
                        {
                            created.Shapes.Placeholders[2]
                                .TextFrame.TextRange.Text =
                                string.Join("\r", lines);
                        }
                        catch
                        {
                        }
                    }
                }

                added++;
            }

            return "Appended " + added +
                " marked draft slides at the end of the " +
                "presentation. Nothing was saved.";
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

                var bullets = new List<string>();
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

                            bullets.Add(TextBoundary.SingleLine(
                                Convert.ToString(bullet),
                                MaxBulletCharacters));
                        }
                    }
                }

                slides.Add(new DraftSlide(title, bullets));
            }

            return slides;
        }
    }
}
