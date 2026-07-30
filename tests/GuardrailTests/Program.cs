using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using OutlookLocalAIChat;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;

namespace GuardrailTests
{
    internal static class Program
    {
        private static int _passed;

        private static int Main()
        {
            try
            {
                Run("HTTPS endpoint is accepted", HttpsEndpointIsAccepted);
                Run("Loopback HTTP endpoint is accepted", LoopbackHttpIsAccepted);
                Run("Remote HTTP endpoint is rejected", RemoteHttpIsRejected);
                Run("Text boundary removes controls and truncates", TextIsBounded);
                Run("Request exposes no model tools", RequestHasNoTools);
                Run("Request schema is messages only", RequestSchemaIsMessagesOnly);
                Run("Email is labeled as untrusted data", EmailIsUntrustedData);
                Run("Conversation history is bounded", HistoryIsBounded);
                Run("Draft service exposes no send capability", DraftHasNoSend);
                Run(
                    "Office startup COM interfaces are dual",
                    OfficeStartupInterfacesAreDual);
                Console.WriteLine("PASS: " + _passed + " guardrail tests");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine("FAIL: " + exception.Message);
                return 1;
            }
        }

        private static void Run(string name, Action test)
        {
            test();
            _passed++;
            Console.WriteLine("PASS: " + name);
        }

        private static void HttpsEndpointIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "https://ai.example.test/v1",
                    out endpoint),
                "HTTPS should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/chat/completions",
                "The chat completions path was not normalized.");
        }

        private static void LoopbackHttpIsAccepted()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://127.0.0.1:1234/v1",
                    out endpoint),
                "Loopback HTTP should be accepted.");
        }

        private static void RemoteHttpIsRejected()
        {
            Uri endpoint;
            Assert(
                !AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1",
                    out endpoint),
                "Remote HTTP must be rejected.");
        }

        private static void TextIsBounded()
        {
            var result = TextBoundary.PlainText("a\u0000bcd", 3);
            Assert(result == "abc", "Unexpected bounded text: " + result);
        }

        private static void RequestHasNoTools()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var json = new JavaScriptSerializer().Serialize(request);
            Assert(!json.Contains("\"tools\""), "Request contains tools.");
            Assert(
                !json.Contains("\"tool_choice\""),
                "Request contains tool_choice.");
            Assert(
                !json.Contains("\"function_call\""),
                "Request contains function_call.");
            Assert(json.Contains("\"stream\":false"), "Streaming must be off.");
        }

        private static void RequestSchemaIsMessagesOnly()
        {
            var fields = typeof(ChatCompletionRequest)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(property => property.Name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[] { "messages", "model", "stream" };
            Assert(
                fields.SequenceEqual(expected),
                "Model request capabilities changed: " +
                string.Join(", ", fields));
        }

        private static void EmailIsUntrustedData()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var context = request.messages[1].content;
            Assert(
                context.Contains("<selected_email>") &&
                context.Contains("untrusted reference data"),
                "Email boundary markers are missing.");
        }

        private static void HistoryIsBounded()
        {
            var history = Enumerable.Range(0, 30)
                .Select(index => new ChatTurn("user", "turn " + index))
                .ToList();
            var request = MakeRequest(history);
            var historyMessages = request.messages.Count - 3;
            Assert(
                historyMessages == TextBoundary.MaxConversationTurns,
                "Unexpected retained history count: " + historyMessages);
        }

        private static void DraftHasNoSend()
        {
            var methods = typeof(DraftService)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method => method.DeclaringType == typeof(DraftService))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Length == 2 &&
                methods.Contains("CreateReplyDraft") &&
                methods.Contains("CreateNewDraft"),
                "Draft service public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void OfficeStartupInterfacesAreDual()
        {
            Assert(
                typeof(IDTExtensibility2).IsImport,
                "IDTExtensibility2 must be a COM-import interface.");
            Assert(
                typeof(IRibbonExtensibility).IsImport,
                "IRibbonExtensibility must be a COM-import interface.");

            AssertDual(typeof(IDTExtensibility2));
            AssertDual(typeof(IRibbonExtensibility));

            var addIn = new AddIn();
            AssertComInterface(addIn, typeof(IDTExtensibility2));
            AssertComInterface(addIn, typeof(IRibbonExtensibility));
        }

        private static void AssertDual(Type interfaceType)
        {
            var attribute = interfaceType
                .GetCustomAttributes(typeof(TypeLibTypeAttribute), false)
                .Cast<TypeLibTypeAttribute>()
                .Single();
            var expected =
                TypeLibTypeFlags.FDispatchable |
                TypeLibTypeFlags.FDual;
            Assert(
                (attribute.Value & expected) == expected,
                interfaceType.Name + " must be a dual dispatch interface.");
        }

        private static void AssertComInterface(object instance, Type interfaceType)
        {
            var pointer = Marshal.GetComInterfaceForObject(
                instance,
                interfaceType);
            try
            {
                Assert(
                    pointer != IntPtr.Zero,
                    interfaceType.Name + " was not exposed by the add-in.");
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    Marshal.Release(pointer);
                }
            }
        }

        private static ChatCompletionRequest MakeRequest(
            IReadOnlyList<ChatTurn> history)
        {
            return ChatRequestFactory.Create(
                "local-model",
                new MessageSnapshot(
                    "entry",
                    "store",
                    "Subject",
                    "Sender",
                    "Recipient",
                    DateTime.UtcNow,
                    "Message body"),
                history,
                "Help me reply.");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }
    }
}
