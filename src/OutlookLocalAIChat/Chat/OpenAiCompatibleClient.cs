using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Outlook;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Chat
{
    public sealed class OpenAiCompatibleClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public OpenAiCompatibleClient()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(120)
            };
        }

        public async Task<string> CompleteAsync(
            AppSettings settings,
            MessageSnapshot message,
            IReadOnlyList<ChatTurn> history,
            string userPrompt,
            CancellationToken cancellationToken)
        {
            if (settings == null || !settings.IsConfigured)
            {
                throw new InvalidOperationException(
                    "Open Settings and configure the endpoint, model, and API key.");
            }

            Uri endpoint;
            if (!AppSettings.TryGetChatCompletionsUri(
                settings.BaseUrl,
                out endpoint))
            {
                throw new InvalidOperationException(
                    "The configured endpoint is invalid.");
            }

            var requestModel = ChatRequestFactory.Create(
                settings.Model,
                message,
                history,
                userPrompt);
            var requestJson = _serializer.Serialize(requestModel);

            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                request.Content = new StringContent(
                    requestJson,
                    Encoding.UTF8,
                    "application/json");

                using (var response = await _httpClient
                    .SendAsync(
                        request,
                        HttpCompletionOption.ResponseHeadersRead,
                        cancellationToken)
                    .ConfigureAwait(true))
                {
                    var responseText = await ReadBoundedAsync(
                        response.Content,
                        cancellationToken).ConfigureAwait(true);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "The AI endpoint returned HTTP " +
                            (int)response.StatusCode +
                            ". Check the endpoint, model, and API key.");
                    }

                    ChatCompletionResponse completion;
                    try
                    {
                        completion =
                            _serializer.Deserialize<ChatCompletionResponse>(
                                responseText);
                    }
                    catch (InvalidOperationException exception)
                    {
                        throw new InvalidOperationException(
                            "The AI endpoint returned invalid JSON.",
                            exception);
                    }

                    var content =
                        completion?.choices != null &&
                        completion.choices.Count > 0
                            ? completion.choices[0]?.message?.content
                            : null;

                    if (string.IsNullOrWhiteSpace(content))
                    {
                        throw new InvalidOperationException(
                            "The AI endpoint returned no message text.");
                    }

                    return TextBoundary.PlainText(
                        content,
                        TextBoundary.MaxAssistantCharacters);
                }
            }
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }

        private static async Task<string> ReadBoundedAsync(
            HttpContent content,
            CancellationToken cancellationToken)
        {
            if (content.Headers.ContentLength.HasValue &&
                content.Headers.ContentLength.Value >
                TextBoundary.MaxHttpResponseCharacters)
            {
                throw new InvalidOperationException(
                    "The AI endpoint response was too large.");
            }

            using (var stream = await content.ReadAsStreamAsync()
                .ConfigureAwait(true))
            using (var reader = new StreamReader(
                stream,
                Encoding.UTF8,
                true,
                4096,
                false))
            {
                var builder = new StringBuilder();
                var buffer = new char[4096];
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var read = await reader.ReadAsync(
                        buffer,
                        0,
                        buffer.Length).ConfigureAwait(true);
                    if (read == 0)
                    {
                        break;
                    }

                    if (builder.Length + read >
                        TextBoundary.MaxHttpResponseCharacters)
                    {
                        throw new InvalidOperationException(
                            "The AI endpoint response was too large.");
                    }

                    builder.Append(buffer, 0, read);
                }

                return builder.ToString();
            }
        }
    }
}
