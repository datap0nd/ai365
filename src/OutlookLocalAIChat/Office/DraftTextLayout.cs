using System;
using System.Collections.Generic;

namespace OutlookLocalAIChat.Office
{
    // Turns the model's plain draft text into a bounded list of
    // styled paragraphs for the Word draft writer: headings, list
    // items, and inline bold segments. This is a local, structural
    // parse - the text itself is never interpreted as HTML or code,
    // and unknown markers simply stay visible as plain text.
    public static class DraftTextLayout
    {
        public const int MaxBoldRangesPerParagraph = 20;

        public const int KindNormal = 0;
        public const int KindHeading1 = 1;
        public const int KindHeading2 = 2;
        public const int KindHeading3 = 3;
        public const int KindBullet = 4;
        public const int KindNumbered = 5;

        public sealed class Paragraph
        {
            internal Paragraph(
                string text,
                int kind,
                IReadOnlyList<BoldRange> boldRanges)
            {
                Text = text ?? string.Empty;
                Kind = kind;
                BoldRanges = boldRanges ?? new BoldRange[0];
            }

            public string Text { get; }

            public int Kind { get; }

            public IReadOnlyList<BoldRange> BoldRanges { get; }
        }

        public sealed class BoldRange
        {
            internal BoldRange(int start, int length)
            {
                Start = start;
                Length = length;
            }

            public int Start { get; }

            public int Length { get; }
        }

        public static IReadOnlyList<Paragraph> Parse(string body)
        {
            var paragraphs = new List<Paragraph>();
            if (body == null)
            {
                return paragraphs;
            }

            var lines = body
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split('\n');
            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                var kind = KindNormal;
                var content = line;
                if (line.StartsWith("### ", StringComparison.Ordinal))
                {
                    kind = KindHeading3;
                    content = line.Substring(4);
                }
                else if (line.StartsWith("## ", StringComparison.Ordinal))
                {
                    kind = KindHeading2;
                    content = line.Substring(3);
                }
                else if (line.StartsWith("# ", StringComparison.Ordinal))
                {
                    kind = KindHeading1;
                    content = line.Substring(2);
                }
                else if (
                    line.StartsWith("- ", StringComparison.Ordinal) ||
                    line.StartsWith("* ", StringComparison.Ordinal))
                {
                    kind = KindBullet;
                    content = line.Substring(2);
                }
                else if (line == "---")
                {
                    content = string.Empty;
                }
                else
                {
                    var prefix = NumberedPrefixLength(line);
                    if (prefix > 0)
                    {
                        kind = KindNumbered;
                        content = line.Substring(prefix);
                    }
                }

                paragraphs.Add(BuildParagraph(content, kind));
            }

            return paragraphs;
        }

        private static int NumberedPrefixLength(string line)
        {
            var index = 0;
            while (index < line.Length &&
                   char.IsDigit(line[index]))
            {
                index++;
            }

            return index > 0 && index + 1 < line.Length &&
                   line[index] == '.' && line[index + 1] == ' '
                ? index + 2
                : 0;
        }

        // Strips ** pairs and records the bold spans they wrapped.
        // An unmatched ** stays visible as literal text.
        private static Paragraph BuildParagraph(
            string content,
            int kind)
        {
            if (content.IndexOf(
                    "**",
                    StringComparison.Ordinal) < 0)
            {
                return new Paragraph(content, kind, null);
            }

            var text = new System.Text.StringBuilder(
                content.Length);
            var ranges = new List<BoldRange>();
            var position = 0;
            while (position < content.Length)
            {
                var open = content.IndexOf(
                    "**",
                    position,
                    StringComparison.Ordinal);
                if (open < 0 ||
                    ranges.Count == MaxBoldRangesPerParagraph)
                {
                    text.Append(content.Substring(position));
                    break;
                }

                var close = content.IndexOf(
                    "**",
                    open + 2,
                    StringComparison.Ordinal);
                if (close < 0)
                {
                    text.Append(content.Substring(position));
                    break;
                }

                text.Append(content.Substring(
                    position,
                    open - position));
                var boldText = content.Substring(
                    open + 2,
                    close - open - 2);
                if (boldText.Length > 0)
                {
                    ranges.Add(new BoldRange(
                        text.Length,
                        boldText.Length));
                }

                text.Append(boldText);
                position = close + 2;
            }

            return new Paragraph(text.ToString(), kind, ranges);
        }
    }
}
