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
                IReadOnlyList<DraftBullet> bullets)
            {
                Title = title ?? string.Empty;
                Bullets = bullets ?? new DraftBullet[0];
            }

            internal string Title { get; }

            internal IReadOnlyList<DraftBullet> Bullets { get; }
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

                slides.Add(new DraftSlide(title, bullets));
            }

            return slides;
        }
    }
}
