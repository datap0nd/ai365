using System;
using System.Text;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Configuration
{
    // One user-configured MCP server. The user adds these
    // deliberately in Settings; nothing in email, document, or model
    // text can register a server. A server is either a local command
    // (stdio transport) or an HTTP(S) endpoint (streamable HTTP
    // transport), and runs entirely with the user's own Windows
    // permissions, outside the add-in's guardrails - which is why
    // the settings page carries an explicit trust notice.
    public sealed class McpServerConfig
    {
        public const int MaxServers = 8;

        public string Name { get; set; } = string.Empty;

        // Local executable path or HTTP(S) endpoint URL.
        public string Target { get; set; } = string.Empty;

        // Raw command-line arguments for stdio servers.
        public string Arguments { get; set; } = string.Empty;

        public bool Enabled { get; set; } = true;

        public bool IsHttp
        {
            get
            {
                return Target.StartsWith(
                           "http://",
                           StringComparison.OrdinalIgnoreCase) ||
                       Target.StartsWith(
                           "https://",
                           StringComparison.OrdinalIgnoreCase);
            }
        }

        public McpServerConfig Sanitized()
        {
            return new McpServerConfig
            {
                Name = SanitizeName(Name),
                Target = TextBoundary.SingleLine(Target, 400),
                Arguments = TextBoundary.SingleLine(
                    Arguments,
                    1000),
                Enabled = Enabled
            };
        }

        // Server names become part of tool names, so they are
        // reduced to a short lowercase token.
        public static string SanitizeName(string value)
        {
            var bounded = TextBoundary.SingleLine(value, 24)
                .ToLowerInvariant();
            var builder = new StringBuilder(bounded.Length);
            foreach (var character in bounded)
            {
                if ((character >= 'a' && character <= 'z') ||
                    (character >= '0' && character <= '9') ||
                    character == '_')
                {
                    builder.Append(character);
                }
                else if (character == '-' || character == ' ')
                {
                    builder.Append('_');
                }
            }

            var name = builder.ToString().Trim('_');
            return name.Length > 0 ? name : "server";
        }
    }
}
