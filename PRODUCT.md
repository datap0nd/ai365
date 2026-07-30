# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows

## Stack

Delegated: C# on .NET Framework 4.8 as a classic Outlook COM add-in, with a
Windows Forms chat window and a single-file Windows installer. This targets
Microsoft Office Professional Plus 2021 on Windows without Microsoft 365 add-in
deployment.

## Users

The primary user works in classic Outlook on a managed Windows work PC and wants
to discuss the selected email, refine a response through conversation, and open
an unsent draft for final review.

## Product Purpose

The add-in provides a small local chat surface beside Outlook. It sends a bounded
plain-text snapshot of the selected message and the current in-memory conversation
to a user-configured OpenAI-compatible endpoint. It can then create an Outlook
reply draft or blank-addressed new-message draft from text the user explicitly
chooses.

Success means installation is understandable, configuration takes one endpoint,
model name, and API key, and no model response can invoke an Outlook send action.

## Positioning

Model output has no executable capability. The model client receives immutable
message text and returns bounded plain text. Only user interface events can call a
separate Outlook draft service, and that service exposes create, save, and display
operations but no send operation.

## Operating Context

- Microsoft Office Professional Plus 2021 with classic Outlook on Windows.
- Per-user local installation is preferred.
- The user selects or opens a message, presses an AI Chat ribbon button, and works
  in a separate compact chat window.
- Configuration is stored for the current Windows user. The API key is encrypted
  with Windows Data Protection API.
- Conversations are kept in memory and disappear when the chat window closes.

## Capabilities and Constraints

- Read only the currently selected or open Outlook mail item.
- Hold a text conversation about that message.
- Generate text suitable for a reply or a new message.
- Create and display an unsent Outlook draft only after an explicit user click.
- Never send, schedule, move, delete, mark, categorize, or modify the source email.
- Never expose tools or function calls to the model endpoint.
- Never render model output as HTML or execute it as code.
- Support an OpenAI-compatible `/v1/chat/completions` endpoint.
- Permit HTTPS endpoints and loopback HTTP endpoints for local model servers.
- Target both 32-bit and 64-bit Office from one installer when practical.
- A production installer should be code-signed. Signing credentials are not
  included in the repository.

## Brand Commitments

The product name is Outlook Local AI Chat. The interface should feel like a
restrained Windows productivity utility, not an AI showcase. Language must be
direct, calm, and explicit about what data is read and when a draft is created.

## Evidence on Hand

The product brief is the user's requested workflow. There are no approved company
logos, claims, screenshots, or signing certificates, and future work must not
fabricate them.

## Product Principles

- Capabilities, not prompts, define the security boundary.
- Nothing leaves the selected-message and active-conversation scope.
- Drafting always ends in Outlook's normal editor with the user in control.
- Local configuration should be inspectable, reversible, and per-user.
- Familiar Windows behavior is more important than decorative novelty.

## Accessibility & Inclusion

The chat window must support keyboard-only operation, visible focus, system text
scaling, high-contrast-compatible colors, and plain-language error recovery.
