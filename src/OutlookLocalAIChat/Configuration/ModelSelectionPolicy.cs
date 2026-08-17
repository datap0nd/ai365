using System;

namespace OutlookLocalAIChat.Configuration
{
    public static class ModelSelectionPolicy
    {
        public static bool IsGenerativeModel(string model)
        {
            var value = (model ?? string.Empty).Trim();
            return value.Length > 0 &&
                   value.IndexOf(
                       "embedding",
                       StringComparison.OrdinalIgnoreCase) < 0;
        }

        public static string DescriptionFor(string model)
        {
            var value = (model ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return "Choose a model from your endpoint after checking it, or enter one manually.";
            }

            return "The model must support OpenAI-compatible chat tool calls.";
        }
    }
}
