using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookLocalAIChat.Configuration
{
    public sealed class ModelGuideEntry
    {
        internal ModelGuideEntry(
            string exampleId,
            string[] matchTokens,
            string speed,
            string quality,
            bool readsEmailImages,
            string notes)
        {
            ExampleId = exampleId ?? string.Empty;
            MatchTokens = matchTokens ?? new string[0];
            Speed = speed ?? string.Empty;
            Quality = quality ?? string.Empty;
            ReadsEmailImages = readsEmailImages;
            Notes = notes ?? string.Empty;
        }

        public string ExampleId { get; }

        public string Speed { get; }

        public string Quality { get; }

        public bool ReadsEmailImages { get; }

        public string Notes { get; }

        internal string[] MatchTokens { get; }

        public string SummaryLine
        {
            get
            {
                var imageNote = ReadsEmailImages
                    ? "Reads email images (vision)"
                    : "Spreadsheet attachments only; images stay as text metadata";
                return ExampleId +
                    " · " + Speed +
                    " · " + Quality +
                    " · " + imageNote;
            }
        }
    }

    public static class ModelCatalog
    {
        private static readonly ModelGuideEntry[] MasterList =
        {
            new ModelGuideEntry(
                "qwen3-vl-30b",
                new[] { "qwen", "vl" },
                "Medium",
                "Very good",
                true,
                "Only listed model that receives email images through MailAI vision input. Best for screenshots, scans, and inline photos."),
            new ModelGuideEntry(
                "qwen3.6-35b-a3b",
                new[] { "qwen", "35", "a3b" },
                "Medium",
                "Very good",
                false,
                "Strong mailbox tool use and drafting. Good default for Excel/CSV attachments and everyday mail questions."),
            new ModelGuideEntry(
                "qwen3.8-27b",
                new[] { "qwen", "27" },
                "Medium",
                "Very good",
                false,
                "Balanced text model when the 35B route feels heavy. Handles spreadsheets; not for image understanding."),
            new ModelGuideEntry(
                "gemma-4-31b-it",
                new[] { "gemma", "31" },
                "Medium",
                "Very good",
                false,
                "Higher-quality Gemma option for longer answers and careful drafting."),
            new ModelGuideEntry(
                "gemma-4-26b-a4b-it",
                new[] { "gemma", "26", "a4b" },
                "Fast",
                "Good",
                false,
                "Lightweight Gemma route for quick mailbox questions and spreadsheet attachments."),
            new ModelGuideEntry(
                "gpt-oss-20b",
                new[] { "gpt", "oss", "20" },
                "Fast",
                "Good",
                false,
                "Fastest general option for simple searches, summaries, and spreadsheet attachments."),
            new ModelGuideEntry(
                "gpt-oss-120b",
                new[] { "gpt", "oss", "120" },
                "Slow",
                "Best",
                false,
                "Highest-quality text model here, but expect much longer waits on large mail context.")
        };

        public static IReadOnlyList<ModelGuideEntry> GuideEntries
        {
            get { return MasterList; }
        }

        public static bool IsDisallowedModel(string model)
        {
            return (model ?? string.Empty).IndexOf(
                "gauss",
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool SupportsVision(string model)
        {
            return Resolve(model)?.ReadsEmailImages ?? false;
        }

        public static ModelGuideEntry Resolve(string model)
        {
            var normalized = Normalize(model);
            if (normalized.Length == 0)
            {
                return null;
            }

            ModelGuideEntry best = null;
            var bestScore = 0;
            foreach (var entry in MasterList)
            {
                var score = MatchScore(normalized, entry.MatchTokens);
                if (score > bestScore)
                {
                    bestScore = score;
                    best = entry;
                }
            }

            return bestScore > 0 ? best : null;
        }

        public static string DescribeForSelection(string model)
        {
            var value = (model ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return BuildGuideOverview();
            }

            if (IsDisallowedModel(value))
            {
                return "This model is hidden because MailAI excludes Gauss and Gausso variants.";
            }

            var profile = Resolve(value);
            if (profile == null)
            {
                return "Custom model. It must support OpenAI-compatible chat tool calls. " +
                    "Only vision models can interpret email images; spreadsheets are sent as extracted text for any model.";
            }

            var builder = new StringBuilder();
            builder.Append(profile.SummaryLine);
            if (profile.Notes.Length > 0)
            {
                builder.Append("\n");
                builder.Append(profile.Notes);
            }

            return builder.ToString();
        }

        public static string FindMatchingDiscoveredId(
            ModelGuideEntry entry,
            IEnumerable<string> discoveredModels)
        {
            if (entry == null)
            {
                return string.Empty;
            }

            foreach (var model in discoveredModels ??
                Enumerable.Empty<string>())
            {
                if (Resolve(model) == entry)
                {
                    return model;
                }
            }

            return string.Empty;
        }

        public static string BuildGuideOverview()
        {
            var lines = new List<string>
            {
                "Model guide (Refresh models loads IDs from your endpoint):"
            };
            foreach (var entry in MasterList)
            {
                lines.Add("• " + entry.SummaryLine);
            }

            lines.Add(
                "Gauss and Gausso models are removed automatically. " +
                "Use qwen3-vl-30b when the email includes images you need interpreted.");
            return string.Join("\n", lines);
        }

        private static int MatchScore(
            string normalizedModel,
            string[] tokens)
        {
            if (tokens == null || tokens.Length == 0)
            {
                return 0;
            }

            var score = 0;
            foreach (var token in tokens)
            {
                var normalizedToken = Normalize(token);
                if (normalizedToken.Length == 0)
                {
                    continue;
                }

                if (normalizedModel.IndexOf(
                        normalizedToken,
                        StringComparison.Ordinal) >= 0)
                {
                    score++;
                }
            }

            return score == tokens.Length ? score : 0;
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace(" ", string.Empty)
                .Replace("_", "-");
        }
    }
}
