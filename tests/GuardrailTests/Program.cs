using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookLocalAIChat;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Interop;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.UI;
using OutlookLocalAIChat.Utilities;

namespace GuardrailTests
{
    internal static class Program
    {
        private static int _passed;

        private static int Main()
        {
            try
            {
                Run(
                    "Vision-capable models are detected broadly",
                    VisionCapableModelsAreDetectedBroadly);
                Run(
                    "Vision prefetch injects image input",
                    VisionPrefetchInjectsImageInput);
                Run(
                    "Vision auto-switch picks the best discovered model",
                    VisionAutoSwitchPicksBestDiscoveredModel);
                Run(
                    "Gauss models are excluded from discovery",
                    GaussModelsAreExcluded);
                Run(
                    "Vision models receive multimodal image follow-up",
                    VisionModelsReceiveMultimodalFollowUp);
                Run(
                    "System prompt states image capability",
                    SystemPromptStatesImageCapability);
                Run(
                    "Vision image limits are enforced",
                    VisionImageLimitsAreEnforced);
                Run(
                    "Web-hosted body images are disclosed as unreadable",
                    WebHostedImagesAreDisclosed);
                Run(
                    "Extensionless pasted images are sniffed and read",
                    PastedImagesWithoutExtensionAreRead);
                Run(
                    "Oversized images are downscaled for vision input",
                    OversizedImagesAreDownscaledForVision);
                Run(
                    "Model catalog describes vision capability",
                    ModelCatalogDescribesVisionCapability);
                Run("HTTPS endpoint is accepted", HttpsEndpointIsAccepted);
                Run("Loopback HTTP endpoint is accepted", LoopbackHttpIsAccepted);
                Run("Remote HTTP endpoint is rejected", RemoteHttpIsRejected);
                Run(
                    "Remote HTTP endpoint requires explicit opt in",
                    RemoteHttpIsAcceptedWithOptIn);
                Run(
                    "Models endpoint is normalized",
                    ModelsEndpointIsNormalized);
                Run(
                    "Default model is empty on install",
                    DefaultModelIsEmptyOnInstall);
                Run(
                    "Email attachments are bounded and readable",
                    EmailAttachmentsAreBounded);
                Run(
                    "Compatible endpoint model discovery is verified",
                    ModelDiscoveryUsesCompatibleContract);
                Run(
                    "Endpoint check uses a lightweight tool probe",
                    EndpointCheckUsesLightweightProbe);
                Run(
                    "Compatible tool calls tolerate a missing call id",
                    ToolCallResponseIsNormalized);
                Run("Text boundary removes controls and truncates", TextIsBounded);
                Run(
                    "Model emphasis becomes native formatting",
                    ModelEmphasisIsNormalized);
                Run(
                    "Endpoint diagnostics expose transport details",
                    EndpointDiagnosticsExposeTransportDetails);
                Run("Mailbox tools are read only", MailboxToolsAreReadOnly);
                Run(
                    "Local search command is explicit and bounded",
                    LocalSearchCommandIsBounded);
                Run(
                    "Working set exposes only ten approved emails",
                    WorkingSetIsStrictlyBounded);
                Run(
                    "Outlook multi-selection accepts one to ten emails",
                    OutlookMultiSelectionIsBounded);
                Run(
                    "Active Explorer selection is used for Send to MetoMail",
                    ActiveExplorerSelectionIsUsed);
                Run(
                    "External context is explicit and bounded",
                    ExternalContextIsBounded);
                Run(
                    "Writing profile analysis is consent-bound and editable",
                    ToneProfileIsBounded);
                Run(
                    "Latest user prompt locally gates draft tools",
                    DraftToolsRequireLocalIntentAuthorization);
                Run(
                    "Draft authorization creates at most one unsent draft",
                    DraftAuthorizationCreatesOnlyOneDraft);
                Run(
                    "Reply draft uses the exact retrieved message handle",
                    ReplyDraftUsesExactHandle);
                Run(
                    "Reply draft rejects missing or fabricated handles",
                    ReplyDraftRequiresIssuedHandle);
                Run(
                    "Linked draft updates the same visible Outlook item",
                    LinkedDraftUpdatesSameItem);
                Run(
                    "Draft HTML is encoded and locally formatted",
                    DraftHtmlIsSafe);
                Run(
                    "Mixed draft tool calls do not consume permission",
                    MixedDraftToolCallIsRejected);
                Run("Request schema exposes bounded tools", RequestSchemaIsBounded);
                Run("Email is labeled as untrusted data", EmailIsUntrustedData);
                Run("Conversation history is bounded", HistoryIsBounded);
                Run("Draft host exposes no send capability", DraftHasNoSend);
                Run(
                    "Compiled add-in contains no Outlook send invocation",
                    CompiledAddInHasNoSendInvocation);
                Run(
                    "Mailbox host exposes one guarded dispatcher",
                    MailboxHostHasGuardedDispatcher);
                Run(
                    "Draft host exposes one guarded dispatcher",
                    DraftHostHasGuardedDispatcher);
                Run(
                    "Office startup and task pane COM interfaces are dual",
                    OfficeStartupInterfacesAreDual);
                Run(
                    "Chat pane is a registered COM control",
                    ChatPaneIsComControl);
                Run(
                    "Outlook ribbon includes Send to MetoMail",
                    RibbonIncludesSendToAi);
                Run(
                    "Selected subjects hide reply and forward prefixes",
                    SelectedSubjectIsCleaned);
                Run(
                    "Self update is official, silent, and restarts Outlook",
                    SelfUpdateIsOfficialAndBounded);
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

        private static void RemoteHttpIsAcceptedWithOptIn()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetChatCompletionsUri(
                    "http://ai.example.test/v1/chat/completions",
                    true,
                    out endpoint),
                "Explicitly allowed remote HTTP should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "http://ai.example.test/v1/chat/completions",
                "Remote HTTP endpoint was not normalized correctly.");

            var settings = new AppSettings
            {
                BaseUrl = "http://ai.example.test/v1",
                Model = "local-model",
                ApiKey = "test-key",
                AllowInsecureHttp = true
            };
            Assert(
                settings.IsConfigured,
                "Remote HTTP opt in was not honored by configuration.");
        }

        private static void ModelsEndpointIsNormalized()
        {
            Uri endpoint;
            Assert(
                AppSettings.TryGetModelsUri(
                    "https://ai.example.test/v1/chat/completions",
                    false,
                    out endpoint),
                "Models endpoint should be accepted.");
            Assert(
                endpoint.AbsoluteUri ==
                "https://ai.example.test/v1/models",
                "The models path was not normalized.");
        }

        private static void DefaultModelIsEmptyOnInstall()
        {
            var settings = new AppSettings();
            Assert(
                settings.Model == string.Empty &&
                !settings.IsConfigured,
                "A fresh install should start without a configured model.");
        }

        private static void EmailAttachmentsAreBounded()
        {
            Assert(
                EmailAttachmentReader.MaxAttachments == 10,
                "Email attachments should allow up to ten files.");

            var csvPath = Path.Combine(
                Path.GetTempPath(),
                "MetoMail-test-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(
                csvPath,
                "Name,Amount\nWidget,42\nGadget,17");
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "budget.csv",
                    csvPath));
                var mail = new FakeSelectedMailItem(
                    "attachment-entry",
                    "Quarterly report")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "attachment-entry",
                    "store",
                    mail);
                var reader = new MessageReader(application);
                var snapshot = reader.CaptureById(
                    "attachment-entry",
                    "store");
                Assert(
                    snapshot.AttachmentNames.Count == 1 &&
                    snapshot.AttachmentNames[0] == "budget.csv",
                    "Supported attachment names were not captured.");

                var request = ChatRequestFactory.Create(
                    "local-model",
                    snapshot,
                    new List<ChatTurn>(),
                    "Help me reply.");
                var reference =
                    MessageContent(request.messages[1]);
                Assert(
                    reference.Contains("Supported attachments (1): budget.csv"),
                    "Attachment metadata was not exposed in the context reference.");

                var host = new MailboxToolHost(
                    application,
                    snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-attachment",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.Content.Contains("\"attachments\"") &&
                    loaded.Content.Contains("budget.csv") &&
                    loaded.Content.Contains("Widget") &&
                    loaded.Content.Contains("Gadget"),
                    "Attachment content was not loaded through read_messages.");
            }
            finally
            {
                if (File.Exists(csvPath))
                {
                    File.Delete(csvPath);
                }
            }
        }

        private static void ModelDiscoveryUsesCompatibleContract()
        {
            const string response =
                "{\"data\":[" +
                "{\"id\":\"qwen3.5-35b-a3b\"}," +
                "{\"id\":\"gpt-oss-20b\"}," +
                "{\"id\":\"text-embedding-model\"}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var models = client.GetModelsAsync(
                    settings,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine == "GET /v1/models HTTP/1.1",
                    "Unexpected model-list request: " +
                    server.RequestLine);
                Assert(
                    server.Authorization == "Bearer test-key",
                    "The API key was not sent as a Bearer token.");
                Assert(
                    models.SequenceEqual(
                        new[]
                        {
                            "gpt-oss-20b",
                            "qwen3.5-35b-a3b"
                        }),
                    "Unexpected discovered models: " +
                    string.Join(", ", models));
            }
        }

        private static void EndpointCheckUsesLightweightProbe()
        {
            var probe = ChatRequestFactory.CreateEndpointCheck(
                "local-model");
            var json = new JavaScriptSerializer()
                .Serialize(probe);
            Assert(
                probe.tools.Count == 1 &&
                probe.tools[0].function.name ==
                MailboxToolCatalog.SearchMailbox &&
                probe.max_tokens.HasValue &&
                probe.max_tokens.Value >= 64 &&
                probe.max_tokens.Value <= 256 &&
                json.Contains("\"tool_choice\"") &&
                json.Contains("search_mailbox") &&
                !json.Contains("read_messages") &&
                !json.Contains("read_thread"),
                "The endpoint check probe was not minimized.");
        }

        private static void ToolCallResponseIsNormalized()
        {
            const string response =
                "{\"choices\":[{\"message\":{" +
                "\"role\":\"assistant\",\"content\":null," +
                "\"tool_calls\":[{\"function\":{" +
                "\"name\":\"search_mailbox\"," +
                "\"arguments\":\"{}\"}}]}}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = EndpointSettings(server.BaseUrl);
                var request = MakeRequest(
                    new List<ChatTurn>());
                request.model = settings.Model;
                var message = client.CompleteAsync(
                    settings,
                    request,
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    server.RequestLine ==
                    "POST /v1/chat/completions HTTP/1.1",
                    "Unexpected completion request: " +
                    server.RequestLine);
                Assert(
                    server.Body.Contains(
                        "\"model\":\"local-model\"") &&
                    server.Body.Contains("\"stream\":false") &&
                    server.Body.Contains("\"tools\""),
                    "The completion request contract is incomplete.");
                Assert(
                    message.tool_calls.Count == 1 &&
                    message.tool_calls[0].id == "call_1" &&
                    message.tool_calls[0].type == "function",
                    "The missing tool-call identity was not normalized.");
            }
        }

        private static AppSettings EndpointSettings(
            string baseUrl)
        {
            return new AppSettings
            {
                BaseUrl = baseUrl,
                Model = "local-model",
                ApiKey = "test-key"
            };
        }

        private static void TextIsBounded()
        {
            var result = TextBoundary.PlainText("a\u0000bcd", 3);
            Assert(result == "abc", "Unexpected bounded text: " + result);
        }

        private static void ModelEmphasisIsNormalized()
        {
            var formatted = SafeModelText.Format(
                "**Decision** and *deadline*\n" +
                "* First action\n" +
                "Unmatched * marker and 2*3",
                TextBoundary.MaxAssistantCharacters);
            Assert(
                formatted.PlainText ==
                    "Decision and deadline\n" +
                    "- First action\n" +
                    "Unmatched  marker and 2*3" &&
                formatted.BoldRanges.Count == 2 &&
                formatted.BoldRanges[0].Start == 0 &&
                formatted.BoldRanges[0].Length == 8 &&
                formatted.BoldRanges[1].Start == 13 &&
                formatted.BoldRanges[1].Length == 8,
                "Model formatting markers were not safely normalized.");
        }

        private static void EndpointDiagnosticsExposeTransportDetails()
        {
            var exception = new AiEndpointException(
                "NETWORK_CONNECT_FAILURE",
                "The endpoint could not be reached.",
                transportDetails:
                    "SocketError ConnectionRefused NativeError 10061");
            var diagnostic = exception.ToDiagnosticText();
            Assert(
                diagnostic.Contains("[NETWORK_CONNECT_FAILURE]") &&
                diagnostic.Contains("Transport details:") &&
                diagnostic.Contains("ConnectionRefused") &&
                diagnostic.Contains("10061"),
                "Transport diagnostics are incomplete: " + diagnostic);
        }

        private static void MailboxToolsAreReadOnly()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var json = new JavaScriptSerializer().Serialize(request);
            Assert(
                json.Contains("\"tools\"") &&
                json.Contains("\"tool_choice\":\"auto\"") &&
                json.Contains("\"maximum\":10") &&
                json.Contains("\"maxItems\":10"),
                "Request does not expose bounded mailbox tools.");
            Assert(json.Contains("\"stream\":false"), "Streaming must be off.");

            var names = request.tools
                .Select(tool => tool.function.name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "read_messages",
                "read_thread",
                "search_mailbox"
            };
            Assert(
                names.SequenceEqual(expected),
                "Unexpected mailbox tools: " +
                string.Join(", ", names));
        }

        private static void LocalSearchCommandIsBounded()
        {
            var search = LocalSearchCommand.Parse(
                "/search project topic");
            var help = LocalSearchCommand.Parse("/search");
            var clear = LocalSearchCommand.Parse(
                "/SEARCH   clear");
            var ordinary = LocalSearchCommand.Parse(
                "/searching project topic");
            var longSearch = LocalSearchCommand.Parse(
                "/search " + new string('x', 500));

            Assert(
                search.Kind == LocalSearchCommandKind.Search &&
                search.Query == "project topic" &&
                help.Kind == LocalSearchCommandKind.Help &&
                clear.Kind == LocalSearchCommandKind.Clear &&
                ordinary.Kind == LocalSearchCommandKind.None &&
                longSearch.Query.Length == 240,
                "The local /search command contract is incomplete.");
        }

        private static void WorkingSetIsStrictlyBounded()
        {
            var messages = Enumerable.Range(1, 12)
                .Select(index => new MessageSnapshot(
                    "entry-" + index,
                    "store",
                    "Subject " + index,
                    "Sender " + index,
                    "Recipient",
                    DateTime.UtcNow.AddMinutes(-index),
                    "private-body-" + index))
                .ToList();
            var normalized = MailboxWorkingSet.Normalize(
                messages.Concat(new[] { messages[0] }));
            Assert(
                normalized.Count == 10 &&
                MailboxWorkingSet.MaxMessages == 10 &&
                MailboxWorkingSet.HandleAt(0) == "context1" &&
                MailboxWorkingSet.HandleAt(9) == "context10",
                "The working-set normalization exceeded ten unique emails.");

            var request = MakeRequest(
                new List<ChatTurn>(),
                workingMessages: normalized);
            var names = request.tools
                .Select(tool => tool.function.name)
                .ToArray();
            var reference =
                MessageContent(request.messages[1]);
            var system =
                MessageContent(request.messages[0]);
            Assert(
                names.SequenceEqual(new[] { "read_messages" }) &&
                reference.Contains("<working_email_set") &&
                reference.Contains("context1") &&
                reference.Contains("context10") &&
                !reference.Contains("private-body-") &&
                system.Contains("working set") &&
                system.Contains("Do not search"),
                "A locked working set exposed search, threads, or email bodies.");

            var host = new MailboxToolHost(
                new FakeOutlookApplication(),
                null,
                normalized);
            var loaded = host.Execute(
                MailboxCall(
                    "read-working-set",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"context1\",\"context2\"," +
                    "\"context3\",\"context4\",\"context5\"," +
                    "\"context6\",\"context7\",\"context8\"," +
                    "\"context9\",\"context10\"]}"));
            var duplicate = host.Execute(
                MailboxCall(
                    "read-duplicate",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"context1\"]}"));
            var lockedSearch = host.Execute(
                MailboxCall(
                    "locked-search",
                    MailboxToolCatalog.SearchMailbox,
                    "{\"query\":\"anything\"}"));
            var lockedThread = host.Execute(
                MailboxCall(
                    "locked-thread",
                    MailboxToolCatalog.ReadThread,
                    "{\"handle\":\"context1\"}"));
            Assert(
                loaded.StatusText.Contains("Request total: 10 of 10") &&
                loaded.Content.Contains("private-body-1") &&
                loaded.Content.Contains("private-body-10") &&
                duplicate.Content.Contains("already_loaded") &&
                lockedSearch.Content.Contains(
                    "MAILBOX_WORKING_SET_LOCKED") &&
                lockedThread.Content.Contains(
                    "MAILBOX_WORKING_SET_LOCKED"),
                "The working-set host bypassed its ten-email lock.");
        }

        private static void OutlookMultiSelectionIsBounded()
        {
            var selection = new FakeSelection(
                Enumerable.Range(1, 3)
                    .Select(index =>
                        (object)new FakeSelectedMailItem(
                            "entry-" + index,
                            "Subject " + index))
                    .ToArray());
            var reader = new MessageReader(new object());
            var messages = reader.CaptureSelectionMany(
                new FakeExplorerContext(selection));
            Assert(
                messages.Count == 3 &&
                messages[0].EntryId == "entry-1" &&
                messages[2].EntryId == "entry-3",
                "Ctrl+click selection did not preserve the selected emails.");

            var overflow = false;
            try
            {
                reader.CaptureSelectionMany(
                    new FakeSelection(
                        new object[MailboxWorkingSet.MaxMessages + 1]));
            }
            catch (InvalidOperationException exception)
            {
                overflow = exception.Message.Contains(
                    "no more than ten");
            }

            Assert(
                overflow,
                "A selection larger than ten emails was not rejected.");
        }

        private static void ActiveExplorerSelectionIsUsed()
        {
            var selection = new FakeSelection(
                new object[]
                {
                    new FakeSelectedMailItem(
                        "active-entry",
                        "Active subject")
                });
            var application = new FakeOutlookApplication
            {
                Explorer = new FakeExplorerContext(selection)
            };
            var messages = new MessageReader(application)
                .CaptureActiveSelectionMany();
            Assert(
                messages.Count == 1 &&
                messages[0].EntryId == "active-entry" &&
                messages[0].Subject == "Active subject",
                "The active Outlook Explorer selection was not captured.");
        }

        private static void ExternalContextIsBounded()
        {
            var documents = Enumerable.Range(1, 5)
                .Select(index => new ExternalContextDocument(
                    "context-" + index + ".txt",
                    new string(
                        (char)('a' + index),
                        9000)))
                .ToList();
            var normalized = ExternalContextDocument.Normalize(documents);
            var total = normalized.Sum(document =>
                document.Content.Length);
            var request = MakeRequest(
                new List<ChatTurn>(),
                externalContext: documents);
            var reference =
                MessageContent(request.messages[1]);
            Assert(
                normalized.Count == ExternalContextDocument.MaxDocuments &&
                total <= ExternalContextDocument.MaxTotalCharacters &&
                reference.Contains("<external_context count=\"3\"") &&
                reference.Contains("untrusted reference data") &&
                !reference.Contains("context-4.txt"),
                "External context exceeded its document or text boundary.");
        }

        private static void ToneProfileIsBounded()
        {
            var samples = Enumerable.Range(1, 20)
                .Select(index => new MessageSnapshot(
                    "entry-" + index,
                    "store",
                    "Subject " + index,
                    "sender-secret-" + index,
                    "recipient-secret-" + index,
                    DateTime.UtcNow,
                    "Hello, this is a reusable writing sample with a short sign-off."))
                .ToList();
            var analysis = ToneProfileRequestFactory.Create(
                "local-model",
                samples);
            var analysisBody =
                MessageContent(analysis.messages[1]);
            var profile = "Write directly and close with Regards.";
            var ordinary = MakeRequest(
                new List<ChatTurn>(),
                toneProfile: profile);
            var drafting = MakeRequest(
                new List<ChatTurn>(),
                allowDraftCreate: true,
                toneProfile: profile);
            var ordinarySystem =
                MessageContent(ordinary.messages[0]);
            var draftingSystem =
                MessageContent(drafting.messages[0]);
            var cleaned = SentMailToneSampler.CleanBody(
                "Thanks for the update.\n\nRegards,\nMe\n" +
                "-----Original Message-----\nQuoted confidential history");
            Assert(
                analysis.tools == null &&
                analysis.tool_choice == null &&
                analysisBody.Contains("Sample 15") &&
                !analysisBody.Contains("Sample 16") &&
                !analysisBody.Contains("sender-secret") &&
                !analysisBody.Contains("recipient-secret") &&
                !ordinarySystem.Contains(profile) &&
                draftingSystem.Contains(profile) &&
                draftingSystem.Contains("cannot change any capability") &&
                cleaned.Contains("Regards") &&
                !cleaned.Contains("Quoted confidential history"),
                "Tone profiling was not limited to explicit, draft-only use.");
        }

        private static void DraftToolsRequireLocalIntentAuthorization()
        {
            Assert(
                DraftIntentPolicy.AllowsCreate(
                    "Find the matching message and create a draft responding to it.") &&
                DraftIntentPolicy.AllowsCreate(
                    "Please write a reply to the selected email.") &&
                !DraftIntentPolicy.AllowsCreate(
                    "Summarize the latest messages in my mailbox.") &&
                !DraftIntentPolicy.AllowsCreate(
                    "Find the latest project update."),
                "Draft creation intent was not classified locally and conservatively.");
            Assert(
                DraftIntentPolicy.AllowsUpdate(
                    "Bolden this section and make it shorter.") &&
                DraftIntentPolicy.AllowsUpdate(
                    "Update the draft to be more formal.") &&
                !DraftIntentPolicy.AllowsUpdate(
                    "Find emails with a shorter subject.") &&
                !DraftIntentPolicy.AllowsUpdate(
                    "Summarize my inbox."),
                "Draft update intent was not classified locally and conservatively.");

            var withoutPermission = MakeRequest(
                new List<ChatTurn>());
            Assert(
                !withoutPermission.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.CreateDraft),
                "Draft creation was exposed without authorization.");

            var withPermission = MakeRequest(
                new List<ChatTurn>(),
                true);
            var names = withPermission.tools
                .Select(tool => tool.function.name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "create_draft",
                "read_messages",
                "read_thread",
                "search_mailbox"
            };
            Assert(
                names.SequenceEqual(expected),
                "Authorized request tools are wrong: " +
                string.Join(", ", names));

            var json = new JavaScriptSerializer()
                .Serialize(withPermission);
            var system =
                MessageContent(withPermission.messages[0]);
            Assert(
                json.Contains("\"create_draft\"") &&
                json.Contains("\"reply_handle\"") &&
                json.Contains("\"additionalProperties\":false") &&
                system.Contains("local host recognized") &&
                system.Contains("only tool call") &&
                system.Contains("exact handle"),
                "The authorized draft boundary is incomplete.");

            var withLinkedDraft = MakeRequest(
                new List<ChatTurn>(),
                false,
                new DraftReference(
                    "new",
                    "Draft subject",
                    "recipient@example.test",
                    string.Empty,
                    "Current body"),
                true);
            Assert(
                withLinkedDraft.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.UpdateDraft) &&
                !withLinkedDraft.tools.Any(tool =>
                    tool.function.name ==
                    DraftToolCatalog.CreateDraft) &&
                MessageContent(withLinkedDraft.messages[2]).Contains(
                        "<linked_draft_reference>"),
                "A linked draft did not replace create with the bounded update tool.");

            var linkedWithoutUpdateIntent = MakeRequest(
                new List<ChatTurn>(),
                false,
                new DraftReference(
                    "new",
                    "Draft subject",
                    "recipient@example.test",
                    string.Empty,
                    "Current body"),
                false);
            Assert(
                !linkedWithoutUpdateIntent.tools.Any(tool =>
                    DraftToolCatalog.IsDraftTool(
                        tool.function.name)) &&
                !linkedWithoutUpdateIntent.messages
                    .OfType<ChatCompletionInputMessage>()
                    .Any(message => MessageContent(message).Contains(
                        "<linked_draft_reference>")),
                "A linked draft was exposed without local update intent.");

            var application = new FakeOutlookApplication();
            var unauthorizedHost = new DraftToolHost(
                application);
            var rejected = unauthorizedHost.Execute(
                DraftCall(
                    "unauthorized",
                    "{\"kind\":\"new\",\"body\":\"Blocked\"}"),
                null,
                new OneShotDraftAuthorization(false),
                true);
            Assert(
                rejected.Content.Contains(
                    "DRAFT_PERMISSION_NOT_AVAILABLE") &&
                application.CreatedCount == 0,
                "A fabricated draft tool call bypassed local authorization.");
        }

        private static void DraftAuthorizationCreatesOnlyOneDraft()
        {
            var application = new FakeOutlookApplication();
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(
                application);

            var first = host.Execute(
                DraftCall(
                    "draft-1",
                    "{\"kind\":\"new\"," +
                    "\"body\":\"Hello\"," +
                    "\"subject\":\"**Subject**\\nInjected\"," +
                    "\"to\":\"one@example.test\\ntwo@example.test\"}"),
                null,
                authorization,
                true);
            var second = host.Execute(
                DraftCall(
                    "draft-2",
                    "{\"kind\":\"new\",\"body\":\"Second\"}"),
                null,
                new OneShotDraftAuthorization(true),
                true);

            Assert(
                authorization.IsConsumed &&
                authorization.IsCreated &&
                application.CreatedCount == 1,
                "The one-shot permission created more than one draft.");
            Assert(
                first.Content.Contains("\"sent\":false") &&
                second.Content.Contains(
                    "DRAFT_ALREADY_LINKED"),
                "The one-shot result contract is incomplete.");
            Assert(
                application.LastDraft.Subject ==
                    "Subject Injected" &&
                application.LastDraft.To ==
                    "one@example.test two@example.test" &&
                application.LastDraft.HTMLBody.Contains("Hello") &&
                application.LastDraft.Saved &&
                application.LastDraft.Displayed &&
                !application.LastDraft.DisplayModal,
                "The unsent draft fields or lifecycle are wrong.");
        }

        private static void LinkedDraftUpdatesSameItem()
        {
            var application = new FakeOutlookApplication();
            var host = new DraftToolHost(application);
            var createAuthorization =
                new OneShotDraftAuthorization(true);
            host.Execute(
                DraftCall(
                    "draft-create",
                    "{\"kind\":\"new\",\"body\":\"First version\"}"),
                null,
                createAuthorization,
                true);

            var original = application.LastDraft;
            var updateAuthorization =
                new OneShotDraftAuthorization(false, true);
            var updated = host.Execute(
                DraftCall(
                    "draft-update",
                    "{\"body\":\"Final section\"," +
                    "\"bold_phrases\":[\"Final\"]}",
                    DraftToolCatalog.UpdateDraft),
                null,
                updateAuthorization,
                true);

            Assert(
                updateAuthorization.IsUpdated &&
                application.CreatedCount == 1 &&
                ReferenceEquals(original, application.LastDraft) &&
                application.LastDraft.HTMLBody.Contains(
                    "<strong>Final</strong> section") &&
                application.LastDraft.SaveCount == 2 &&
                application.LastDraft.DisplayCount == 2 &&
                updated.Content.Contains("\"action\":\"updated\""),
                "The live update did not mutate and redisplay the same draft.");

            var secondUpdate = host.Execute(
                DraftCall(
                    "draft-update-2",
                    "{\"body\":\"Should not apply\"}",
                    DraftToolCatalog.UpdateDraft),
                null,
                updateAuthorization,
                true);
            Assert(
                secondUpdate.Content.Contains(
                    "DRAFT_UPDATE_NOT_AVAILABLE") &&
                application.LastDraft.SaveCount == 2,
                "One request updated the linked draft more than once.");
        }

        private static void ReplyDraftUsesExactHandle()
        {
            var application = new FakeOutlookApplication();
            var wrong = application.RegisterReplySource(
                "wrong-entry",
                "store",
                "RE: Wrong latest message",
                "wrong.sender@example.test");
            var target = application.RegisterReplySource(
                "target-entry",
                "store",
                "RE: Target project update",
                "target.sender@example.test");
            var wrongSnapshot = new MessageSnapshot(
                "wrong-entry",
                "store",
                "Wrong latest message",
                "wrong.sender@example.test",
                "recipient@example.test",
                DateTime.UtcNow,
                "Wrong body");
            var targetSnapshot = new MessageSnapshot(
                "target-entry",
                "store",
                "Target project update",
                "target.sender@example.test",
                "recipient@example.test",
                DateTime.UtcNow.AddMinutes(-5),
                "Target body");
            Func<string, MessageSnapshot> resolver = handle =>
                handle == "m2"
                    ? targetSnapshot
                    : handle == "selected"
                        ? wrongSnapshot
                        : null;
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(application);

            var result = host.Execute(
                DraftCall(
                    "reply-target",
                    "{\"kind\":\"reply\"," +
                    "\"reply_handle\":\"m2\"," +
                    "\"body\":\"Hello **Target contact**\"}"),
                resolver,
                authorization,
                true);

            Assert(
                authorization.IsCreated &&
                result.Content.Contains("\"draft_kind\":\"reply\"") &&
                target.ReplyCount == 1 &&
                wrong.ReplyCount == 0 &&
                application.LastDraft.To ==
                    "target.sender@example.test" &&
                application.LastDraft.Subject ==
                    "RE: Target project update" &&
                application.LastDraft.HTMLBody.Contains(
                    "Hello <strong>Target contact</strong>") &&
                !application.LastDraft.HTMLBody.Contains("**") &&
                host.ActiveDraft.Body == "Hello Target contact",
                "The reply was not bound to the exact retrieved handle.");
        }

        private static void ReplyDraftRequiresIssuedHandle()
        {
            var application = new FakeOutlookApplication();
            var source = application.RegisterReplySource(
                "target-entry",
                "store",
                "RE: Target",
                "target@example.test");
            var snapshot = new MessageSnapshot(
                "target-entry",
                "store",
                "Target",
                "target@example.test",
                "recipient@example.test",
                DateTime.UtcNow,
                "Body");
            Func<string, MessageSnapshot> resolver = handle =>
                handle == "selected" ? snapshot : null;

            var missingAuthorization =
                new OneShotDraftAuthorization(true);
            var missing = new DraftToolHost(application).Execute(
                DraftCall(
                    "reply-missing",
                    "{\"kind\":\"reply\",\"body\":\"Hello\"}"),
                resolver,
                missingAuthorization,
                true);
            var unknownAuthorization =
                new OneShotDraftAuthorization(true);
            var unknown = new DraftToolHost(application).Execute(
                DraftCall(
                    "reply-unknown",
                    "{\"kind\":\"reply\"," +
                    "\"reply_handle\":\"fabricated-id\"," +
                    "\"body\":\"Hello\"}"),
                resolver,
                unknownAuthorization,
                true);

            Assert(
                missing.Content.Contains(
                    "DRAFT_REPLY_HANDLE_REQUIRED") &&
                unknown.Content.Contains(
                    "DRAFT_REPLY_HANDLE_UNKNOWN") &&
                !missingAuthorization.IsConsumed &&
                !unknownAuthorization.IsConsumed &&
                source.ReplyCount == 0,
                "A missing or fabricated reply handle reached Outlook.");
        }

        private static void DraftHtmlIsSafe()
        {
            var html = SafeDraftHtml.Format(
                "Hello <script>alert('x')</script>\nImportant",
                new[] { "Important" });
            Assert(
                !html.Contains("<script>") &&
                html.Contains("&lt;script&gt;") &&
                html.Contains("<div style=") &&
                html.Contains("<strong>Important</strong>"),
                "Draft HTML did not encode untrusted markup: " + html);

            var markdown = SafeDraftHtml.FormatContent(
                "Hello **Target contact**\n* Next step\n__Important__\n" +
                "*Deadline* and stray * marker",
                new string[0]);
            Assert(
                markdown.PlainText ==
                    "Hello Target contact\n- Next step\nImportant\n" +
                    "Deadline and stray  marker" &&
                markdown.Html.Contains(
                    "Hello <strong>Target contact</strong>") &&
                markdown.Html.Contains(
                    "<strong>Important</strong>") &&
                markdown.Html.Contains(
                    "<strong>Deadline</strong>") &&
                !markdown.Html.Contains("**") &&
                !markdown.Html.Contains("__") &&
                !markdown.Html.Contains("<script>"),
                "Markdown notation was not converted to safe email formatting: " +
                markdown.Html);

            var visual = SafeDraftHtml.FormatContent(
                "# Project update\n## Decisions\n- First item\n- Second item\n" +
                "1. Confirm owner\n2. Confirm date\n---\nNext step",
                new[] { "Confirm owner" });
            Assert(
                visual.Html.Contains("<h2") &&
                visual.Html.Contains("<h3") &&
                visual.Html.Contains("<ul") &&
                visual.Html.Contains("<ol") &&
                visual.Html.Contains("<hr") &&
                visual.Html.Contains("<strong>Confirm owner</strong>") &&
                !visual.Html.Contains("<script>") &&
                !visual.Html.Contains("<img"),
                "Safe visual email layout was not rendered locally: " +
                visual.Html);

            var application = new FakeOutlookApplication();
            var rejected = new DraftToolHost(application).Execute(
                DraftCall(
                    "html-injection",
                    "{\"kind\":\"new\",\"body\":\"Safe\"," +
                    "\"html\":\"<img src=x>\"}"),
                null,
                new OneShotDraftAuthorization(true),
                true);
            Assert(
                rejected.Content.Contains(
                    "DRAFT_ARGUMENTS_INVALID") &&
                application.CreatedCount == 0,
                "Arbitrary model HTML reached the Outlook draft path.");
        }

        private static void MixedDraftToolCallIsRejected()
        {
            var application = new FakeOutlookApplication();
            var authorization =
                new OneShotDraftAuthorization(true);
            var host = new DraftToolHost(
                application);
            var result = host.Execute(
                DraftCall(
                    "mixed",
                    "{\"kind\":\"new\",\"body\":\"Hello\"}"),
                null,
                authorization,
                false);

            Assert(
                result.Content.Contains(
                    "DRAFT_TOOL_MUST_BE_EXCLUSIVE") &&
                !authorization.IsConsumed &&
                application.CreatedCount == 0,
                "A mixed draft call bypassed exclusivity.");
        }

        private static ChatToolCall DraftCall(
            string id,
            string arguments,
            string name = DraftToolCatalog.CreateDraft)
        {
            return new ChatToolCall
            {
                id = id,
                type = "function",
                function = new ChatToolCallFunction
                {
                    name = name,
                    arguments = arguments
                }
            };
        }

        private static ChatToolCall MailboxCall(
            string id,
            string name,
            string arguments)
        {
            return new ChatToolCall
            {
                id = id,
                type = "function",
                function = new ChatToolCallFunction
                {
                    name = name,
                    arguments = arguments
                }
            };
        }

        private static void RequestSchemaIsBounded()
        {
            var fields = typeof(ChatCompletionRequest)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(property => property.Name)
                .OrderBy(name => name)
                .ToArray();
            var expected = new[]
            {
                "max_tokens",
                "messages",
                "model",
                "stream",
                "tool_choice",
                "tools"
            };
            Assert(
                fields.SequenceEqual(expected),
                "Model request capabilities changed: " +
                string.Join(", ", fields));
        }

        private static void EmailIsUntrustedData()
        {
            var request = MakeRequest(new List<ChatTurn>());
            var context =
                MessageContent(request.messages[1]);
            Assert(
                context.Contains("<selected_email_reference") &&
                context.Contains("untrusted reference data") &&
                !context.Contains("Message body"),
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
            var methods = typeof(DraftToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method => method.DeclaringType == typeof(DraftToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                !methods.Any(name =>
                    name.IndexOf("Send", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Delete", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Move", StringComparison.OrdinalIgnoreCase) >= 0),
                "Draft host exposes a forbidden capability: " +
                string.Join(", ", methods));
        }

        private static void CompiledAddInHasNoSendInvocation()
        {
            var forbidden = new HashSet<string>(
                new[]
                {
                    "Send",
                    "SendAndReceive",
                    "Submit"
                },
                StringComparer.Ordinal);
            var violations = new List<string>();
            foreach (var type in typeof(AddIn).Assembly.GetTypes())
            {
                var methods = type
                    .GetMethods(
                        BindingFlags.Instance |
                        BindingFlags.Static |
                        BindingFlags.Public |
                        BindingFlags.NonPublic)
                    .Cast<MethodBase>()
                    .Concat(type.GetConstructors(
                        BindingFlags.Instance |
                        BindingFlags.Static |
                        BindingFlags.Public |
                        BindingFlags.NonPublic))
                    .Concat(new MethodBase[]
                    {
                        type.TypeInitializer
                    }.Where(item => item != null));
                foreach (var method in methods)
                {
                    foreach (var value in LoadedStrings(method))
                    {
                        if (forbidden.Contains(value))
                        {
                            violations.Add(
                                type.FullName + "." +
                                method.Name + ":" + value);
                        }
                    }
                }
            }

            Assert(
                violations.Count == 0,
                "Compiled Outlook send invocation found: " +
                string.Join(", ", violations));
        }

        private static IEnumerable<string> LoadedStrings(
            MethodBase method)
        {
            MethodBody body;
            try
            {
                body = method.GetMethodBody();
            }
            catch
            {
                body = null;
            }

            var il = body == null
                ? null
                : body.GetILAsByteArray();
            if (il == null)
            {
                yield break;
            }

            var oneByte = new OpCode[256];
            var twoByte = new OpCode[256];
            foreach (var field in typeof(OpCodes).GetFields(
                BindingFlags.Public | BindingFlags.Static))
            {
                var opcode = (OpCode)field.GetValue(null);
                var value = (ushort)opcode.Value;
                if (value < 0x100)
                {
                    oneByte[value] = opcode;
                }
                else if ((value & 0xff00) == 0xfe00)
                {
                    twoByte[value & 0xff] = opcode;
                }
            }

            var position = 0;
            while (position < il.Length)
            {
                var first = il[position++];
                var opcode = first == 0xfe
                    ? twoByte[il[position++]]
                    : oneByte[first];
                if (opcode.OperandType == OperandType.InlineString)
                {
                    var token = BitConverter.ToInt32(
                        il,
                        position);
                    position += 4;
                    string value;
                    try
                    {
                        value = method.Module.ResolveString(token);
                    }
                    catch
                    {
                        continue;
                    }

                    yield return value;
                    continue;
                }

                position += OperandSize(
                    opcode.OperandType,
                    il,
                    position);
            }
        }

        private static int OperandSize(
            OperandType operandType,
            byte[] il,
            int position)
        {
            switch (operandType)
            {
                case OperandType.InlineNone:
                    return 0;
                case OperandType.ShortInlineBrTarget:
                case OperandType.ShortInlineI:
                case OperandType.ShortInlineVar:
                    return 1;
                case OperandType.InlineVar:
                    return 2;
                case OperandType.InlineI:
                case OperandType.InlineBrTarget:
                case OperandType.InlineField:
                case OperandType.InlineMethod:
                case OperandType.InlineSig:
                case OperandType.InlineString:
                case OperandType.InlineTok:
                case OperandType.InlineType:
                case OperandType.ShortInlineR:
                    return 4;
                case OperandType.InlineI8:
                case OperandType.InlineR:
                    return 8;
                case OperandType.InlineSwitch:
                    return 4 +
                        BitConverter.ToInt32(il, position) * 4;
                default:
                    throw new InvalidOperationException(
                        "Unknown IL operand type: " + operandType);
            }
        }

        private static void MailboxHostHasGuardedDispatcher()
        {
            var methods = typeof(MailboxToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method =>
                    method.DeclaringType ==
                    typeof(MailboxToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Length == 1 &&
                methods[0] == "Execute",
                "Mailbox tool host public capabilities changed: " +
                string.Join(", ", methods));
        }

        private static void DraftHostHasGuardedDispatcher()
        {
            var methods = typeof(DraftToolHost)
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(method =>
                    method.DeclaringType ==
                    typeof(DraftToolHost))
                .Select(method => method.Name)
                .ToArray();
            Assert(
                methods.Contains("Execute") &&
                methods.Contains("Dispose") &&
                !methods.Contains("Send"),
                "Draft tool host public capabilities changed: " +
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
            Assert(
                typeof(ICustomTaskPaneConsumer).IsImport,
                "ICustomTaskPaneConsumer must be a COM-import interface.");

            AssertDual(typeof(IDTExtensibility2));
            AssertDual(typeof(IRibbonExtensibility));
            AssertDual(typeof(ICustomTaskPaneConsumer));

            var addIn = new AddIn();
            AssertComInterface(addIn, typeof(IDTExtensibility2));
            AssertComInterface(addIn, typeof(IRibbonExtensibility));
            AssertComInterface(
                addIn,
                typeof(ICustomTaskPaneConsumer));
        }

        private static void ChatPaneIsComControl()
        {
            var type = typeof(ChatPane);
            var visible = type
                .GetCustomAttributes(
                    typeof(ComVisibleAttribute),
                    false)
                .Cast<ComVisibleAttribute>()
                .Single();
            var progId = type
                .GetCustomAttributes(
                    typeof(ProgIdAttribute),
                    false)
                .Cast<ProgIdAttribute>()
                .Single();
            Assert(visible.Value, "ChatPane must be COM visible.");
            Assert(
                progId.Value ==
                "OutlookLocalAIChat.ChatPane",
                "Unexpected ChatPane ProgID.");
            Assert(
                type.GUID ==
                new Guid(
                    "14D24FA1-4342-442F-B68B-B68D7372794C"),
                "Unexpected ChatPane CLSID.");
        }

        private static void RibbonIncludesSendToAi()
        {
            var xml = new AddIn().GetCustomUI(
                "Microsoft.Outlook.Explorer");
            Assert(
                xml.Contains("ContextMenuMailItem") &&
                xml.Contains("OnSendToAi") &&
                xml.Contains("Send to MetoMail") &&
                xml.Contains("label=\"MetoMail\""),
                "The Outlook explorer ribbon XML is incomplete: " + xml);
        }

        private static void SelectedSubjectIsCleaned()
        {
            Assert(
                SubjectDisplay.Clean(" RE: FW: Fwd: Quarterly plan ") ==
                    "Quarterly plan" &&
                SubjectDisplay.Clean("Project update") ==
                    "Project update",
                "Selected subject prefixes were not removed safely.");
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
            IReadOnlyList<ChatTurn> history,
            bool allowDraftCreate = false,
            DraftReference activeDraft = null,
            bool allowDraftUpdate = false,
            IReadOnlyList<MessageSnapshot> workingMessages = null,
            IReadOnlyList<ExternalContextDocument> externalContext = null,
            string toneProfile = null)
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
                "Help me reply.",
                allowDraftCreate,
                activeDraft,
                allowDraftUpdate,
                workingMessages,
                externalContext,
                toneProfile);
        }

        private static string MessageContent(object message)
        {
            return Convert.ToString(
                ((ChatCompletionInputMessage)message).content) ??
                string.Empty;
        }

        private static void VisionCapableModelsAreDetectedBroadly()
        {
            Assert(
                ModelCatalog.IsVisionCapable("qwen3-vl-30b") &&
                ModelCatalog.IsVisionCapable("Qwen3-VL-30B-Instruct") &&
                ModelCatalog.IsVisionCapable("my-vision-model") &&
                ModelCatalog.IsVisionCapable("gemma-4-31b-it") &&
                ModelCatalog.IsVisionCapable("gemma-4-26b-a4b-it") &&
                ModelCatalog.IsVisionCapable("gemma3-27b-it") &&
                ModelCatalog.IsVisionCapable("llava-1.6-13b") &&
                ModelCatalog.IsVisionCapable("MiniCPM-V-2_6") &&
                !ModelCatalog.IsVisionCapable("gpt-oss-20b") &&
                !ModelCatalog.IsVisionCapable("qwen3.6-35b-a3b") &&
                !ModelCatalog.IsVisionCapable("text-embedding-vl"),
                "Vision capability detection is too narrow or too broad.");
        }

        private static void VisionPrefetchInjectsImageInput()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "MetoMail-prefetch-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "scan.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "prefetch-entry",
                    "Invoice scan")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "prefetch-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("prefetch-entry", "store");
                var request = ChatRequestFactory.Create(
                    "qwen3-vl-30b",
                    snapshot,
                    new List<ChatTurn>(),
                    "Summarize the image.");
                var host = new MailboxToolHost(application, snapshot);
                Assert(
                    VisionImagePrefetch.TryInject(
                        request,
                        "qwen3-vl-30b",
                        host,
                        snapshot,
                        null),
                    "Vision prefetch did not inject image context.");
                var body = new JavaScriptSerializer()
                    .Serialize(request);
                Assert(
                    body.Contains("\"type\":\"image_url\"") &&
                    body.Contains("data:image/png;base64,"),
                    "Prefetched request is missing multimodal image content.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void VisionAutoSwitchPicksBestDiscoveredModel()
        {
            var discovered = new[]
            {
                "gpt-oss-20b",
                "qwen3.6-35b-a3b",
                "qwen3-vl-30b"
            };
            Assert(
                ModelCatalog.FindBestVisionModel(discovered) ==
                    "qwen3-vl-30b",
                "The best vision model was not selected from the saved list.");
            Assert(
                ModelCatalog.FindBestVisionModel(
                    new[] { "gemma-4-31b-it", "qwen3-vl-30b" }) ==
                    "qwen3-vl-30b",
                "The dedicated vision model should be preferred over multimodal Gemma.");
            Assert(
                ModelCatalog.FindBestVisionModel(
                    new[] { "gpt-oss-20b", "gemma-4-31b-it" }) ==
                    "gemma-4-31b-it",
                "Multimodal Gemma should be used when no dedicated vision model exists.");

            var settings = new AppSettings
            {
                Model = "gpt-oss-20b",
                SwitchToVisionModelForImages = true,
                DiscoveredModels = new List<string>(discovered)
            };
            var snapshot = new MessageSnapshot(
                "entry",
                "store",
                "Invoice",
                "Sender",
                "Recipient",
                DateTime.UtcNow,
                "Body",
                new[] { "scan.png" });
            var routed = ModelRouting.ResolveForRequest(
                settings,
                ModelRouting.ContextMayIncludeImages(snapshot, null));
            Assert(
                routed == "qwen3-vl-30b" &&
                ModelRouting.IsTemporaryVisionSwitch(settings, routed),
                "Image requests did not temporarily switch to the vision model.");

            settings.SwitchToVisionModelForImages = false;
            Assert(
                ModelRouting.ResolveForRequest(
                    settings,
                    true) == "gpt-oss-20b",
                "The checkbox should gate temporary vision switching.");

            settings.SwitchToVisionModelForImages = true;
            settings.Model = "qwen3-vl-30b";
            Assert(
                ModelRouting.ResolveForRequest(
                    settings,
                    true) == "qwen3-vl-30b" &&
                !ModelRouting.IsTemporaryVisionSwitch(
                    settings,
                    "qwen3-vl-30b"),
                "Vision models should not be replaced when already selected.");
        }

        private static void GaussModelsAreExcluded()
        {
            const string response =
                "{\"data\":[" +
                "{\"id\":\"qwen3-vl-30b\"}," +
                "{\"id\":\"gauss-vision-7b\"}," +
                "{\"id\":\"my-gausso-model\"}," +
                "{\"id\":\"gpt-oss-20b\"}]}";
            using (var server = new FakeEndpoint(response))
            using (var client = new OpenAiCompatibleClient())
            {
                var models = client.GetModelsAsync(
                    EndpointSettings(server.BaseUrl),
                    CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
                server.Wait();

                Assert(
                    models.SequenceEqual(
                        new[]
                        {
                            "gpt-oss-20b",
                            "qwen3-vl-30b"
                        }),
                    "Gauss models were not excluded: " +
                    string.Join(", ", models));
            }

            Assert(
                ModelCatalog.IsDisallowedModel("gauss-vision-7b") &&
                ModelCatalog.IsDisallowedModel("my-gausso-model") &&
                !ModelCatalog.IsDisallowedModel("qwen3-vl-30b"),
                "The Gauss filter did not classify model names correctly.");
        }

        private static void VisionModelsReceiveMultimodalFollowUp()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "MetoMail-test-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "scan.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "vision-entry",
                    "Invoice scan")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "vision-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("vision-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-vision",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].FileName == "scan.png" &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/png;base64,") &&
                    loaded.Content.Contains("vision_available") &&
                    !loaded.Content.Contains("base64,"),
                    "Vision image payloads were not separated from tool JSON.");

                var request = ChatRequestFactory.Create(
                    "qwen3-vl-30b",
                    snapshot,
                    new List<ChatTurn>(),
                    "What is in the attachment?");
                ChatRequestFactory.AppendToolExchange(
                    request,
                    new ChatCompletionResponseMessage
                    {
                        role = "assistant",
                        content = string.Empty,
                        tool_calls = new List<ChatToolCall>
                        {
                            MailboxCall(
                                "read-vision",
                                MailboxToolCatalog.ReadMessages,
                                "{\"handles\":[\"selected\"]}")
                        }
                    },
                    new List<MailboxToolResult> { loaded },
                    "qwen3-vl-30b");
                var body = new JavaScriptSerializer()
                    .Serialize(request);
                Assert(
                    body.Contains("\"type\":\"image_url\"") &&
                    body.Contains("data:image/png;base64,") &&
                    body.Contains("scan.png"),
                    "Vision follow-up did not append multimodal image content.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void SystemPromptStatesImageCapability()
        {
            var snapshot = new MessageSnapshot(
                "entry",
                "store",
                "Invoice",
                "Sender",
                "Recipient",
                DateTime.UtcNow,
                "Body",
                new[] { "scan.png" });

            var textRequest = ChatRequestFactory.Create(
                "gpt-oss-20b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize this image.");
            var textSystem = MessageContent(textRequest.messages[0]);
            Assert(
                textSystem.Contains("text-only") &&
                textSystem.Contains("cannot view images") &&
                textSystem.Contains("Today's date is"),
                "Text-only models are not told about unseen images.");

            var visionRequest = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize this image.");
            var visionSystem = MessageContent(visionRequest.messages[0]);
            Assert(
                visionSystem.Contains("vision-capable model") &&
                visionSystem.Contains("attachment filename"),
                "Vision models are not instructed to use image input.");

            var plainRequest = ChatRequestFactory.Create(
                "gpt-oss-20b",
                new MessageSnapshot(
                    "entry",
                    "store",
                    "Subject",
                    "Sender",
                    "Recipient",
                    DateTime.UtcNow,
                    "Body"),
                new List<ChatTurn>(),
                "Hello.");
            Assert(
                !MessageContent(plainRequest.messages[0])
                    .Contains("text-only"),
                "The image warning should appear only when images are in context.");
        }

        private static void VisionImageLimitsAreEnforced()
        {
            var images = new List<VisionImagePayload>();
            for (var index = 0; index < 12; index++)
            {
                images.Add(new VisionImagePayload(
                    "img" + index + ".png",
                    "data:image/png;base64,AAA" + index));
            }

            images.Add(new VisionImagePayload(
                "huge.png",
                "data:image/png;base64," + new string('A', 2200001)));

            var request = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                null,
                new List<ChatTurn>(),
                "Describe the images.");
            ChatRequestFactory.AppendToolExchange(
                request,
                new ChatCompletionResponseMessage
                {
                    role = "assistant",
                    content = string.Empty,
                    tool_calls = new List<ChatToolCall>
                    {
                        MailboxCall(
                            "read-many",
                            MailboxToolCatalog.ReadMessages,
                            "{\"handles\":[\"selected\"]}")
                    }
                },
                new List<MailboxToolResult>
                {
                    new MailboxToolResult(
                        "read-many",
                        "{}",
                        "loaded",
                        images)
                },
                "qwen3-vl-30b");
            var body = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            }.Serialize(request);
            Assert(
                CountToken(body, "\"type\":\"image_url\"") == 8,
                "The per-request image cap was not applied.");
            Assert(
                !body.Contains("huge.png"),
                "An oversized image data URL was not dropped.");
            Assert(
                body.Contains("omitted"),
                "Omitted images are not disclosed to the model.");
        }

        private static void PastedImagesWithoutExtensionAreRead()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "MetoMail-pasted-" + Guid.NewGuid().ToString("N"));
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJ" +
                    "AAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            try
            {
                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "image001",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "pasted-entry",
                    "Survey request")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "pasted-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("pasted-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-pasted",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/png;base64,"),
                    "An extensionless pasted image was not sniffed as PNG.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void OversizedImagesAreDownscaledForVision()
        {
            var pngPath = Path.Combine(
                Path.GetTempPath(),
                "MetoMail-big-" + Guid.NewGuid().ToString("N") + ".png");
            using (var bitmap = new System.Drawing.Bitmap(
                1400,
                1400,
                System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                var bounds = new System.Drawing.Rectangle(
                    0,
                    0,
                    1400,
                    1400);
                var data = bitmap.LockBits(
                    bounds,
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                var noise = new byte[
                    Math.Abs(data.Stride) * 1400];
                new Random(42).NextBytes(noise);
                Marshal.Copy(
                    noise,
                    0,
                    data.Scan0,
                    noise.Length);
                bitmap.UnlockBits(data);
                bitmap.Save(
                    pngPath,
                    System.Drawing.Imaging.ImageFormat.Png);
            }

            try
            {
                Assert(
                    new FileInfo(pngPath).Length > 1536 * 1024,
                    "The generated test image is too small to exercise downscaling.");

                var attachments = new FakeOutlookAttachments();
                attachments.Add(new FakeOutlookAttachment(
                    "photo.png",
                    pngPath));
                var mail = new FakeSelectedMailItem(
                    "big-image-entry",
                    "Site photo")
                {
                    Attachments = attachments
                };
                var application = new FakeOutlookApplication();
                application.Session.Register(
                    "big-image-entry",
                    "store",
                    mail);
                var snapshot = new MessageReader(application)
                    .CaptureById("big-image-entry", "store");
                var host = new MailboxToolHost(application, snapshot);
                var loaded = host.Execute(
                    MailboxCall(
                        "read-big-image",
                        MailboxToolCatalog.ReadMessages,
                        "{\"handles\":[\"selected\"]}"));
                Assert(
                    loaded.VisionImages.Count == 1 &&
                    loaded.VisionImages[0].DataUrl.StartsWith(
                        "data:image/jpeg;base64,") &&
                    loaded.VisionImages[0].DataUrl.Length <=
                        EmailAttachmentReader.MaxImageDataUrlCharacters &&
                    loaded.Content.Contains("downscaled"),
                    "An oversized image was not downscaled into vision input.");
            }
            finally
            {
                if (File.Exists(pngPath))
                {
                    File.Delete(pngPath);
                }
            }
        }

        private static void WebHostedImagesAreDisclosed()
        {
            var mail = new FakeSelectedMailItem(
                "remote-image-entry",
                "Newsletter")
            {
                HTMLBody =
                    "<html><body>" +
                    "<a href=\"https://example.test/offer\">" +
                    "<img src=\"https://cdn.example.test/banner.jpg\"></a>" +
                    "<img src='https://cdn.example.test/photo.png'>" +
                    "<img src=\"cid:signature-logo\">" +
                    "</body></html>"
            };
            var application = new FakeOutlookApplication();
            application.Session.Register(
                "remote-image-entry",
                "store",
                mail);
            var snapshot = new MessageReader(application)
                .CaptureById("remote-image-entry", "store");
            Assert(
                snapshot.RemoteImageCount == 2,
                "Web-hosted image detection counted " +
                snapshot.RemoteImageCount +
                " instead of 2 (cid images must not count).");

            var request = ChatRequestFactory.Create(
                "qwen3-vl-30b",
                snapshot,
                new List<ChatTurn>(),
                "Summarize the image in this email.");
            var reference = MessageContent(request.messages[1]);
            Assert(
                reference.Contains(
                    "Web-hosted images referenced by URL: 2"),
                "The message reference does not disclose web-hosted images.");

            var host = new MailboxToolHost(application, snapshot);
            var loaded = host.Execute(
                MailboxCall(
                    "read-remote",
                    MailboxToolCatalog.ReadMessages,
                    "{\"handles\":[\"selected\"]}"));
            Assert(
                loaded.Content.Contains(
                    "web_hosted_images_not_included"),
                "The tool result does not disclose web-hosted images.");
            Assert(
                loaded.Content.Contains("web_hosted_images_note") &&
                loaded.Content.Contains("cannot view them"),
                "The tool result is missing the web-hosted image note.");
        }

        private static int CountToken(string text, string token)
        {
            var count = 0;
            var index = 0;
            while ((index = text.IndexOf(
                token,
                index,
                StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += token.Length;
            }

            return count;
        }

        private static void ModelCatalogDescribesVisionCapability()
        {
            Assert(
                ModelCatalog.SupportsVision("qwen3-vl-30b") &&
                ModelCatalog.SupportsVision("gemma-4-31b-it") &&
                !ModelCatalog.SupportsVision("qwen3.6-35b-a3b") &&
                ModelCatalog.GuideEntries.Count >= 7,
                "The model catalog vision flags are incomplete.");

            var description =
                ModelCatalog.DescribeForSelection("qwen3-vl-30b");
            Assert(
                description.Contains("Vision model") &&
                description.Contains("email images"),
                "Vision model guidance is incomplete: " + description);

            var overview = ModelCatalog.BuildGuideOverview();
            Assert(
                overview.Contains("Refresh models") &&
                overview.Length < 80,
                "The model guide overview is incomplete.");
        }

        private static void SelfUpdateIsOfficialAndBounded()
        {
            Assert(
                SelfUpdater.InstallerUrl.StartsWith(
                    "https://github.com/datap0nd/outlook-local-ai-chat/releases/",
                    StringComparison.Ordinal),
                "The updater must download only the official release installer over HTTPS.");

            var script = SelfUpdater.BuildUpdateScript();
            Assert(
                script.Contains("OUTLOOK.EXE") &&
                script.Contains("if %tries% GEQ 150 exit /b 1") &&
                script.Contains(
                    "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART") &&
                script.Contains("start \"\" outlook.exe") &&
                script.Contains("%~1"),
                "The update script must wait for Outlook to close, install " +
                "silently with a bounded wait, and restart Outlook.");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private sealed class FakeEndpoint : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly Task _requestTask;
            private readonly string _responseBody;

            public FakeEndpoint(string responseBody)
            {
                _responseBody = responseBody;
                _listener = new TcpListener(
                    IPAddress.Loopback,
                    0);
                _listener.Start();
                var port =
                    ((IPEndPoint)_listener.LocalEndpoint)
                    .Port;
                BaseUrl = "http://127.0.0.1:" +
                    port + "/v1";
                _requestTask = Task.Run(
                    (Action)HandleRequest);
            }

            public string BaseUrl { get; }

            public string RequestLine { get; private set; } =
                string.Empty;

            public string Authorization { get; private set; } =
                string.Empty;

            public string Body { get; private set; } =
                string.Empty;

            public void Wait()
            {
                if (!_requestTask.Wait(
                    TimeSpan.FromSeconds(10)))
                {
                    throw new InvalidOperationException(
                        "The fake endpoint did not receive a request.");
                }

                if (_requestTask.IsFaulted)
                {
                    throw _requestTask.Exception
                        .GetBaseException();
                }
            }

            public void Dispose()
            {
                _listener.Stop();
                try
                {
                    _requestTask.Wait(
                        TimeSpan.FromSeconds(1));
                }
                catch
                {
                }
            }

            private void HandleRequest()
            {
                using (var client = _listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(
                    stream,
                    Encoding.ASCII,
                    false,
                    4096,
                    true))
                {
                    RequestLine =
                        reader.ReadLine() ?? string.Empty;
                    var contentLength = 0;
                    while (true)
                    {
                        var line = reader.ReadLine();
                        if (string.IsNullOrEmpty(line))
                        {
                            break;
                        }

                        var separator = line.IndexOf(':');
                        if (separator <= 0)
                        {
                            continue;
                        }

                        var name = line
                            .Substring(0, separator)
                            .Trim();
                        var value = line
                            .Substring(separator + 1)
                            .Trim();
                        if (name.Equals(
                            "Authorization",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Authorization = value;
                        }
                        else if (name.Equals(
                            "Content-Length",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            int.TryParse(
                                value,
                                out contentLength);
                        }
                    }

                    if (contentLength > 0)
                    {
                        var buffer =
                            new char[contentLength];
                        var offset = 0;
                        while (offset < contentLength)
                        {
                            var read = reader.Read(
                                buffer,
                                offset,
                                contentLength - offset);
                            if (read <= 0)
                            {
                                break;
                            }

                            offset += read;
                        }

                        Body = new string(
                            buffer,
                            0,
                            offset);
                    }

                    var responseBytes =
                        Encoding.UTF8.GetBytes(
                            _responseBody);
                    var headers = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\n" +
                        "Content-Type: application/json\r\n" +
                        "Content-Length: " +
                        responseBytes.Length +
                        "\r\nConnection: close\r\n\r\n");
                    stream.Write(
                        headers,
                        0,
                        headers.Length);
                    stream.Write(
                        responseBytes,
                        0,
                        responseBytes.Length);
                    stream.Flush();
                }
            }
        }
    }

    public sealed class FakeExplorerContext
    {
        public FakeExplorerContext(FakeSelection selection)
        {
            Selection = selection;
        }

        public FakeSelection Selection { get; }
    }

    public sealed class FakeSelection
    {
        private readonly object[] _items;

        public FakeSelection(object[] items)
        {
            _items = items ?? new object[0];
        }

        public int Count
        {
            get { return _items.Length; }
        }

        public object Item(int index)
        {
            return _items[index - 1];
        }
    }

    public sealed class FakeSelectedMailItem
    {
        public FakeSelectedMailItem(
            string entryId,
            string subject)
        {
            EntryID = entryId;
            Subject = subject;
            Parent = new FakeMailFolder();
            ReceivedTime = DateTime.UtcNow;
            Attachments = new FakeOutlookAttachments();
        }

        public string MessageClass { get; } = "IPM.Note";

        public string EntryID { get; }

        public string Subject { get; }

        public string SenderName { get; } = "Sender";

        public string SenderEmailAddress { get; } =
            "sender@example.test";

        public string To { get; } = "recipient@example.test";

        public string Body { get; } = "Message body";

        public string HTMLBody { get; set; } = string.Empty;

        public DateTime ReceivedTime { get; }

        public DateTime SentOn { get; } = DateTime.MinValue;

        public DateTime CreationTime { get; } = DateTime.MinValue;

        public FakeMailFolder Parent { get; }

        public FakeOutlookAttachments Attachments { get; set; }
    }

    public sealed class FakeOutlookAttachments
    {
        private readonly List<FakeOutlookAttachment> _items =
            new List<FakeOutlookAttachment>();

        public int Count
        {
            get { return _items.Count; }
        }

        public void Add(FakeOutlookAttachment attachment)
        {
            _items.Add(attachment);
        }

        public FakeOutlookAttachment Item(int index)
        {
            return _items[index - 1];
        }
    }

    public sealed class FakeOutlookAttachment
    {
        private readonly string _sourcePath;

        public FakeOutlookAttachment(
            string fileName,
            string sourcePath)
        {
            FileName = fileName;
            _sourcePath = sourcePath;
        }

        public string FileName { get; }

        public void SaveAsFile(string path)
        {
            File.Copy(_sourcePath, path, true);
        }
    }

    public sealed class FakeMailFolder
    {
        public string StoreID { get; } = "store";
    }

    public sealed class FakeOutlookApplication
    {
        private readonly FakeOutlookSession _session =
            new FakeOutlookSession();

        public int CreatedCount { get; private set; }

        public FakeMailItem LastDraft { get; private set; }

        public FakeExplorerContext Explorer { get; set; }

        public FakeOutlookSession Session
        {
            get { return _session; }
        }

        public object ActiveExplorer()
        {
            return Explorer;
        }

        public FakeReplySource RegisterReplySource(
            string entryId,
            string storeId,
            string replySubject,
            string replyTo)
        {
            var source = new FakeReplySource(
                this,
                replySubject,
                replyTo);
            _session.Register(entryId, storeId, source);
            return source;
        }

        public void RecordReply(FakeMailItem draft)
        {
            LastDraft = draft;
        }

        public object CreateItem(int itemType)
        {
            if (itemType != 0)
            {
                throw new InvalidOperationException(
                    "Only mail items are allowed in the test host.");
            }

            CreatedCount++;
            LastDraft = new FakeMailItem();
            return LastDraft;
        }
    }

    public sealed class FakeOutlookSession
    {
        private readonly Dictionary<string, object> _items =
            new Dictionary<string, object>(StringComparer.Ordinal);

        public void Register(
            string entryId,
            string storeId,
            object item)
        {
            _items[Key(entryId, storeId)] = item;
        }

        public object GetItemFromID(string entryId)
        {
            return GetItemFromID(entryId, string.Empty);
        }

        public object GetItemFromID(
            string entryId,
            string storeId)
        {
            object item;
            if (!_items.TryGetValue(
                    Key(entryId, storeId),
                    out item))
            {
                throw new InvalidOperationException(
                    "Unknown fake Outlook item.");
            }

            return item;
        }

        private static string Key(
            string entryId,
            string storeId)
        {
            return (entryId ?? string.Empty) +
                "\n" +
                (storeId ?? string.Empty);
        }
    }

    public sealed class FakeReplySource
    {
        private readonly FakeOutlookApplication _application;
        private readonly string _replySubject;
        private readonly string _replyTo;

        public FakeReplySource(
            FakeOutlookApplication application,
            string replySubject,
            string replyTo)
        {
            _application = application;
            _replySubject = replySubject;
            _replyTo = replyTo;
        }

        public int ReplyCount { get; private set; }

        public object Reply()
        {
            ReplyCount++;
            var draft = new FakeMailItem
            {
                Subject = _replySubject,
                To = _replyTo,
                HTMLBody = "<div>Quoted original</div>"
            };
            _application.RecordReply(draft);
            return draft;
        }
    }

    public sealed class FakeMailItem
    {
        public FakeMailItem()
        {
            Subject = string.Empty;
            To = string.Empty;
            CC = string.Empty;
            HTMLBody = string.Empty;
        }

        public string Subject { get; set; }

        public string To { get; set; }

        public string CC { get; set; }

        public string HTMLBody { get; set; }

        public bool Saved { get; private set; }

        public bool Displayed { get; private set; }

        public bool DisplayModal { get; private set; }

        public int SaveCount { get; private set; }

        public int DisplayCount { get; private set; }

        public void Save()
        {
            Saved = true;
            SaveCount++;
        }

        public void Display(bool modal)
        {
            Displayed = true;
            DisplayModal = modal;
            DisplayCount++;
        }
    }
}
