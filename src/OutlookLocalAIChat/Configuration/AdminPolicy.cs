using System;
using Microsoft.Win32;

namespace OutlookLocalAIChat.Configuration
{
    // Machine/user policy switches for AI365, read from
    // Software\Policies\AI365 in HKLM (admin- or GPO-set, wins) and
    // HKCU. Policies only ever remove capabilities - nothing here
    // can widen what the add-in may do.
    public static class AdminPolicy
    {
        public const string PolicyKeyPath =
            "Software\\Policies\\AI365";

        // DisableGemini = 1 hides and blocks Google Gemini sign-in
        // across the suite; only the user's own OpenAI-compatible
        // endpoint remains available.
        public static bool GeminiDisabled
        {
            get { return ReadFlag("DisableGemini"); }
        }

        private static bool ReadFlag(string valueName)
        {
            return ReadFlag(Registry.LocalMachine, valueName) ||
                   ReadFlag(Registry.CurrentUser, valueName);
        }

        private static bool ReadFlag(
            RegistryKey hive,
            string valueName)
        {
            try
            {
                using (var key = hive.OpenSubKey(PolicyKeyPath))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    var value = key.GetValue(valueName);
                    return value != null &&
                           Convert.ToInt32(value) == 1;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
