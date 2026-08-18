using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.Chat
{
    // Surfaces the user's configured MCP servers as chat tools. Tool
    // names are namespaced mcp_{server}_{tool}, every result is
    // bounded and wrapped as untrusted data, and the exposed tool
    // count is capped. Servers only exist because the user added
    // them in Settings; nothing model-facing can register one, and
    // MCP tools can never reach the add-in's own draft or mailbox
    // capabilities.
    public sealed class McpToolHost : IDisposable
    {
        public const string ToolPrefix = "mcp_";
        public const int MaxExposedTools = 40;

        private readonly List<McpConnection> _connections =
            new List<McpConnection>();
        private readonly Dictionary<string, McpRoute> _routes =
            new Dictionary<string, McpRoute>(
                StringComparer.Ordinal);
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly HttpClient _httpClient;
        private List<ChatToolDefinition> _definitions;

        private sealed class McpRoute
        {
            public McpRoute(
                McpConnection connection,
                string toolName)
            {
                Connection = connection;
                ToolName = toolName;
            }

            public McpConnection Connection { get; }

            public string ToolName { get; }
        }

        public McpToolHost(
            IReadOnlyList<McpServerConfig> servers)
        {
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
            var count = 0;
            foreach (var server in servers ??
                new McpServerConfig[0])
            {
                if (server == null ||
                    !server.Enabled ||
                    server.Target.Trim().Length == 0)
                {
                    continue;
                }

                if (count == McpServerConfig.MaxServers)
                {
                    break;
                }

                _connections.Add(new McpConnection(
                    server.Sanitized(),
                    _httpClient));
                count++;
            }
        }

        public bool HasServers
        {
            get { return _connections.Count > 0; }
        }

        public static bool IsMcpTool(string name)
        {
            return name != null &&
                   name.StartsWith(
                       ToolPrefix,
                       StringComparison.Ordinal);
        }

        // Connects to each enabled server (once per session) and
        // builds the namespaced tool definitions. A server that
        // fails is skipped with a logged error so one bad server
        // never blocks the chat.
        public IReadOnlyList<ChatToolDefinition> GetDefinitions()
        {
            if (_definitions != null)
            {
                return _definitions;
            }

            var definitions = new List<ChatToolDefinition>();
            foreach (var connection in _connections)
            {
                if (definitions.Count >= MaxExposedTools)
                {
                    break;
                }

                IReadOnlyList<McpToolDescriptor> tools;
                try
                {
                    tools = connection.ListTools();
                }
                catch (Exception exception)
                {
                    Log.Error(
                        "McpListTools." + connection.ServerName,
                        exception);
                    continue;
                }

                foreach (var tool in tools)
                {
                    if (definitions.Count >= MaxExposedTools)
                    {
                        break;
                    }

                    var exposedName = BuildExposedName(
                        connection.ServerName,
                        tool.Name);
                    if (_routes.ContainsKey(exposedName))
                    {
                        continue;
                    }

                    _routes[exposedName] = new McpRoute(
                        connection,
                        tool.Name);
                    definitions.Add(new ChatToolDefinition
                    {
                        type = "function",
                        function = new ChatToolFunctionDefinition
                        {
                            name = exposedName,
                            description = TextBoundary.PlainText(
                                "[MCP tool from user-configured " +
                                "server '" +
                                connection.ServerName +
                                "'] " + tool.Description,
                                800),
                            parameters =
                                tool.Schema as
                                    IDictionary<string, object> ??
                                new Dictionary<string, object>
                                {
                                    { "type", "object" },
                                    {
                                        "properties",
                                        new Dictionary<string, object>()
                                    }
                                }
                        }
                    });
                }
            }

            _definitions = definitions;
            return _definitions;
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            var name = call?.function?.name ?? string.Empty;
            McpRoute route;
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id) ||
                !_routes.TryGetValue(name, out route))
            {
                return Error(
                    call?.id,
                    "MCP_TOOL_NOT_ALLOWED",
                    "The requested MCP tool is not registered.");
            }

            try
            {
                IDictionary<string, object> arguments =
                    new Dictionary<string, object>();
                if (!string.IsNullOrWhiteSpace(
                    call.function.arguments))
                {
                    arguments = _serializer.DeserializeObject(
                            call.function.arguments) as
                        IDictionary<string, object> ??
                        new Dictionary<string, object>();
                }

                bool isError;
                var content = route.Connection.CallTool(
                    route.ToolName,
                    arguments,
                    out isError);
                var payload = _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "untrusted_mcp_data", true },
                        { "server", route.Connection.ServerName },
                        { "tool", route.ToolName },
                        { "is_error", isError },
                        { "content", content }
                    });
                if (payload.Length >
                    ContextScale.Scaled(
                        TextBoundary.MaxToolResultCharacters))
                {
                    return Error(
                        call.id,
                        "MCP_RESULT_TOO_LARGE",
                        "The bounded MCP result was still too large to return safely.");
                }

                return new MailboxToolResult(
                    call.id,
                    payload,
                    "MCP " + route.Connection.ServerName + ": " +
                    route.ToolName +
                    (isError ? " reported an error" : " completed"));
            }
            catch (Exception exception)
            {
                Log.Error("McpTool." + name, exception);
                return Error(
                    call.id,
                    "MCP_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "MCP_TOOL_FAILED"));
            }
        }

        public void Dispose()
        {
            foreach (var connection in _connections)
            {
                try
                {
                    connection.Dispose();
                }
                catch
                {
                }
            }

            _connections.Clear();
            _routes.Clear();
            _definitions = null;
            _httpClient.Dispose();
        }

        // mcp_{server}_{tool}, reduced to the characters OpenAI-style
        // endpoints accept and capped at 64.
        private static string BuildExposedName(
            string serverName,
            string toolName)
        {
            var builder = new StringBuilder(
                ToolPrefix + serverName + "_");
            foreach (var character in toolName)
            {
                if ((character >= 'a' && character <= 'z') ||
                    (character >= 'A' && character <= 'Z') ||
                    (character >= '0' && character <= '9') ||
                    character == '_' || character == '-')
                {
                    builder.Append(character);
                }
                else
                {
                    builder.Append('_');
                }

                if (builder.Length >= 64)
                {
                    break;
                }
            }

            return builder.ToString();
        }

        private MailboxToolResult Error(
            string callId,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        }
                    }),
                "[" + code + "] " + message);
        }
    }
}
