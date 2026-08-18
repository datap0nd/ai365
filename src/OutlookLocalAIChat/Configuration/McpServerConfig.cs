using System;
using System.Collections.Generic;
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

        // Optional HTTP headers for HTTP servers, one per line as
        // "Name: value" - typically an Authorization header. Sent
        // only to this server's own endpoint and never logged.
        public string Headers { get; set; } = string.Empty;

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
                Headers = TextBoundary.PlainText(Headers, 2000),
                Enabled = Enabled
            };
        }

        // Parses the Headers text into at most 8 well-formed
        // name/value pairs. Names are restricted to header tokens
        // and reserved protocol headers are dropped so a configured
        // header can never break the MCP transport itself.
        public IReadOnlyList<KeyValuePair<string, string>>
            ParsedHeaders()
        {
            var headers =
                new List<KeyValuePair<string, string>>();
            foreach (var raw in (Headers ?? string.Empty)
                .Replace("\r\n", "\n")
                .Split('\n'))
            {
                if (headers.Count == 8)
                {
                    break;
                }

                var separator = raw.IndexOf(':');
                if (separator <= 0)
                {
                    continue;
                }

                var name = raw.Substring(0, separator).Trim();
                var value = TextBoundary.SingleLine(
                    raw.Substring(separator + 1),
                    500);
                if (name.Length == 0 ||
                    value.Length == 0 ||
                    !IsHeaderToken(name) ||
                    IsReservedHeader(name))
                {
                    continue;
                }

                headers.Add(
                    new KeyValuePair<string, string>(
                        name,
                        value));
            }

            return headers;
        }

        private static bool IsHeaderToken(string name)
        {
            foreach (var character in name)
            {
                if (!(character >= 'a' && character <= 'z') &&
                    !(character >= 'A' && character <= 'Z') &&
                    !(character >= '0' && character <= '9') &&
                    character != '-' &&
                    character != '_')
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsReservedHeader(string name)
        {
            return
                string.Equals(
                    name,
                    "Accept",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    name,
                    "Content-Type",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    name,
                    "Content-Length",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    name,
                    "Host",
                    StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    name,
                    "Mcp-Session-Id",
                    StringComparison.OrdinalIgnoreCase);
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
