using System;

namespace OutlookLocalAIChat.Security
{
    // Local intent gate for the Excel and PowerPoint panes. Document
    // draft surfaces (the AI365 Draft worksheet, appended draft
    // slides, and unsent Outlook email drafts) unlock only when the
    // user's own latest prompt contains an explicit produce-something
    // phrase; model or document text can never grant the permission.
    public static class DocumentDraftIntentPolicy
    {
        private static readonly string[] DraftPhrases =
        {
            "create a draft",
            "create draft",
            "draft sheet",
            "draft slide",
            "draft a slide",
            "draft an email",
            "draft the email",
            "draft this",
            "create a slide",
            "create slides",
            "create a sheet",
            "create a table in",
            "create an email",
            "add a slide",
            "add slides",
            "add a sheet",
            "insert a slide",
            "insert slides",
            "insert a sheet",
            "insert a table",
            "insert into",
            "make a slide",
            "make slides",
            "make a sheet",
            "make a table in",
            "write it to",
            "write this to",
            "write these to",
            "write to the draft",
            "write an email",
            "compose an email",
            "put this in",
            "put it in",
            "put these in",
            "put that in",
            "export to",
            "email this",
            "email these",
            "email it",
            "email that",
            "email the",
            "mail this",
            "mail these",
            "send this to",
            "send these to",
            "send it to",
            "send that to",
            "send to outlook",
            "send to excel",
            "send to powerpoint",
            "prepare a draft",
            "prepare an email",
            "prepare a slide",
            "turn this into"
        };

        public static bool AllowsDraft(string userPrompt)
        {
            var prompt = TextBoundary.PlainText(
                    userPrompt,
                    TextBoundary.MaxUserPromptCharacters)
                .ToLowerInvariant();
            foreach (var phrase in DraftPhrases)
            {
                if (prompt.IndexOf(
                        phrase,
                        StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
