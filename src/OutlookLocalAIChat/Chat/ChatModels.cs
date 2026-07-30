using System.Collections.Generic;

namespace OutlookLocalAIChat.Chat
{
    public sealed class ChatTurn
    {
        public ChatTurn(string role, string content)
        {
            Role = role;
            Content = content;
        }

        public string Role { get; }

        public string Content { get; }
    }

    public sealed class ChatCompletionRequest
    {
        public string model { get; set; }

        public List<ChatCompletionMessage> messages { get; set; }

        public bool stream { get; set; }
    }

    public sealed class ChatCompletionMessage
    {
        public string role { get; set; }

        public string content { get; set; }
    }

    public sealed class ChatCompletionResponse
    {
        public List<ChatCompletionChoice> choices { get; set; }
    }

    public sealed class ChatCompletionChoice
    {
        public ChatCompletionMessage message { get; set; }
    }
}
