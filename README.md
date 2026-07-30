# Outlook Local AI Chat

A Windows-only AI chat add-in for classic Outlook in Microsoft Office
Professional Plus 2021. It installs locally, reads only the selected email, and
opens unsent drafts for human review.

It does not use Microsoft 365 add-in deployment, Microsoft Graph, Entra ID, or
Outlook MCP.

## Install

1. Close Outlook.
2. Download
   [OutlookLocalAIChatSetup.exe](https://github.com/datap0nd/outlook-local-ai-chat/releases/latest/download/OutlookLocalAIChatSetup.exe).
3. Run the installer for your Windows account.
4. Start classic Outlook.
5. Select an email and choose **AI Chat > Open AI Chat** on the ribbon.
6. Open **Settings** and enter:
   - the OpenAI-compatible endpoint or base URL;
   - the model name;
   - the API key.

Examples:

```text
https://ai.example.test/v1
https://ai.example.test/v1/chat/completions
http://127.0.0.1:1234/v1
```

Remote endpoints must use HTTPS. Plain HTTP is accepted only for loopback
addresses such as `localhost` and `127.0.0.1`.

The first unsigned build may trigger a Windows SmartScreen warning. A trusted
code-signing certificate is required to remove that warning for normal company
distribution.

## Use

1. Select or open a normal email.
2. Open **AI Chat**.
3. Ask questions or request draft text.
4. Continue refining the text in the conversation.
5. Choose **Open reply draft** or **Open new draft**.
6. Review, edit, address, and send the message using Outlook's normal editor.

The conversation remains in memory only while the chat window is open.

## Hard security boundary

The model is not an Outlook agent.

- The AI request contains messages only. It contains no tools, functions, or
  executable command schema.
- The model client receives a plain-text snapshot. It never receives the Outlook
  application object or draft service.
- Model output is length-limited plain text displayed in a Windows control. It is
  never evaluated, executed, or rendered as HTML.
- Only explicit local button events can call `CreateReplyDraft` or
  `CreateNewDraft`.
- The draft service exposes no send, move, delete, schedule, or mailbox traversal
  operation.
- Drafts are saved and displayed as unsent Outlook items.
- CI fails if forbidden Outlook action calls are introduced.

These controls prevent model output from reaching an email-send capability in
this implementation. They do not claim protection against a compromised Windows
account, a modified add-in binary, vulnerabilities in Outlook or .NET, or an
administrator replacing the installed files.

See [SECURITY.md](SECURITY.md) for the full threat model.

## Data flow

For every chat request, the add-in sends the configured endpoint:

- selected email subject;
- sender and recipient display strings;
- received timestamp;
- up to 24,000 characters of plain-text body;
- up to 12 recent chat turns;
- the current prompt.

Nothing is sent to Microsoft 365 by the add-in. Outlook itself continues to use
whatever mail server your organization configured.

## API compatibility

The endpoint must support:

```http
POST /v1/chat/completions
Authorization: Bearer YOUR_KEY
Content-Type: application/json
```

The request uses only `model`, `messages`, and `stream: false`. The response must
provide `choices[0].message.content` as text.

## Remove

1. Close Outlook.
2. Open Windows **Installed apps** or **Apps & features**.
3. Uninstall **Outlook Local AI Chat**.

Endpoint settings remain under:

```text
%LOCALAPPDATA%\OutlookLocalAIChat
```

Delete that folder manually if you also want to remove the encrypted API key and
local diagnostic log.

## Build

Requirements:

- Windows 10 or newer
- Visual Studio 2022 Build Tools with .NET Framework 4.8 targeting pack
- Inno Setup 6

Build the assembly and tests:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Restore-StrongNameKey.ps1
msbuild OutlookLocalAIChat.sln /m /p:Configuration=Release
tests\GuardrailTests\bin\Release\GuardrailTests.exe
powershell -ExecutionPolicy Bypass -File scripts\Test-Guardrails.ps1
```

The repository stores the stable strong-name key as Base64 so local and CI builds
use the same COM identity. A strong name is an assembly identity mechanism, not a
trusted publisher signature.

Build the installer:

```powershell
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" `
  "installer\OutlookLocalAIChat.iss"
```

The installer is written to:

```text
artifacts\OutlookLocalAIChatSetup.exe
```

GitHub Actions builds, smoke-tests, and publishes the same single-file installer.

## Compatibility

- Classic Outlook for Windows
- Microsoft Office Professional Plus 2021
- 32-bit or 64-bit Office on Windows
- .NET Framework 4.8

The new Outlook for Windows does not load COM add-ins.
