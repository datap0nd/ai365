---
name: Outlook Local AI Chat
description: A restrained Windows utility for discussing one selected email and opening user-reviewed Outlook drafts.
colors:
  action-blue: "#005fb8"
  window: "Canvas"
  window-text: "CanvasText"
  control: "ButtonFace"
  control-text: "ButtonText"
  highlight: "Highlight"
  highlight-text: "HighlightText"
  secondary-text: "#505050"
  muted-surface: "#f4f6f8"
  border: "ButtonBorder"
  error: "#a32626"
  high-contrast-secondary: "GrayText"
  high-contrast-error: "LinkText"
typography:
  title:
    fontFamily: "system-ui, sans-serif"
    fontSize: "12pt"
    fontWeight: 700
    lineHeight: 1.2
  body:
    fontFamily: "system-ui, sans-serif"
    fontSize: "10pt"
    fontWeight: 400
    lineHeight: 1.4
  label:
    fontFamily: "system-ui, sans-serif"
    fontSize: "9pt"
    fontWeight: 700
    lineHeight: 1.3
  hint:
    fontFamily: "system-ui, sans-serif"
    fontSize: "8pt"
    fontWeight: 400
    lineHeight: 1.3
rounded:
  square: "0px"
spacing:
  tight: "4px"
  compact: "8px"
  control-gap: "10px"
  toolbar-x: "12px"
  content-x: "18px"
  dialog-x: "24px"
components:
  button-primary:
    backgroundColor: "{colors.action-blue}"
    textColor: "{colors.highlight-text}"
    typography: "{typography.label}"
    rounded: "{rounded.square}"
    padding: "0 16px"
    height: "34px"
  button-secondary:
    backgroundColor: "{colors.window}"
    textColor: "{colors.window-text}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "0 14px"
    height: "34px"
  button-link:
    backgroundColor: "{colors.window}"
    textColor: "{colors.action-blue}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "0 4px"
    height: "28px"
  input:
    backgroundColor: "{colors.window}"
    textColor: "{colors.window-text}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "4px 6px"
    height: "34px"
---

# Design System: Outlook Local AI Chat

## Overview

**Creative North Star: "The Guardrailed Desk Tool"**

Outlook Local AI Chat is a compact native Windows productivity utility. It should feel adjacent to classic Outlook, not like an AI showcase: familiar system typography, quiet white and cool-gray surfaces, square fields, restrained density, and blue reserved for direct actions.

The screen tells one ordered story: confirm which email is loaded, discuss that bounded plain-text snapshot, then deliberately open an unsent draft in Outlook. The visual hierarchy must keep this read/chat/draft-only boundary obvious. Model output is conversation content, never an executable command, and no control may imply that the utility can send mail.

**Key Characteristics:**

- Native Windows behavior and system settings take priority over decorative identity.
- One selected-message strip anchors the current scope.
- A large plain-text transcript carries the work without chat bubbles or HTML treatment.
- The composer and AI action are visually bounded together.
- Draft actions are separate, secondary, and enabled only when usable.
- Disclosure and status text state what data is used and what action occurred.

## Colors

The palette is mostly Windows system color roles, with a fixed Outlook-adjacent blue for direct actions in standard contrast and system highlight colors in high contrast.

### Primary

- **Action Blue:** Used for the Send to AI button, transcript speaker emphasis, and toolbar links. In high contrast, use the Windows `Highlight` system role instead.

### Neutral

- **Window:** The main transcript, toolbar, secondary buttons, and draft-action surface use the Windows canvas role.
- **Window Text:** Primary copy follows the Windows canvas text role.
- **Control:** Settings and high-contrast muted bands use the Windows control surface.
- **Control Text:** Settings text follows the Windows control text role.
- **Secondary Text:** Supporting metadata, hints, disclosures, and normal status messages use a quieter neutral in standard contrast and `GrayText` in high contrast.
- **Muted Surface:** Header, composer, and status bands use a cool gray in standard contrast, then fall back to the Windows control surface in high contrast.
- **Border:** Square field and secondary-button outlines use the Windows control-dark or button-border role.
- **Error:** Recoverable errors use restrained dark red in standard contrast and a system-provided high-contrast role when high contrast is active.

### Named Rules

**The System Color Rule.** System canvas, text, control, highlight, and border roles outrank fixed palette values whenever Windows high contrast is active.

**The One Accent Rule.** Blue marks direct actions and the user's transcript label. It is not decoration and must not spread into background panels or assistant content.

**The Text-Plus-State Rule.** Color never carries status alone. Every busy, error, disabled, disclosure, and draft state also has explicit text or native control state.

## Typography

**Display Font:** None

**Body Font:** Windows message-box system font, normally Segoe UI, with the active Windows fallback

**Label Font:** The same system family at bold weight

**Character:** Typography is familiar, compact, and subordinate to the task. It inherits Windows font settings and scales through WinForms `AutoScaleMode.Font`; nominal sizes describe the standard presentation, not a hard override of user settings.

### Hierarchy

- **Title** (bold, nominally 12pt): Selected email subject in the message strip.
- **Body** (regular, nominally 10pt): Transcript turns, composer text, settings fields, and primary reading content.
- **Label** (bold, nominally 9pt): Speaker names, field labels, and the primary action.
- **Hint** (regular, no smaller than 8pt): Keyboard guidance, draft disclosure, metadata, and status copy.

### Named Rules

**The System Font Rule.** Use `SystemFonts.MessageBoxFont` and relative size adjustments. Do not bundle a custom font or pin typography in a way that defeats Windows text scaling.

**The Plain-Text Rule.** Transcript content is plain text only. Do not render model output as HTML, rich cards, markdown actions, links, or executable affordances.

## Layout

The chat window is a single-column, vertically stacked utility sized at 540 by 720 client pixels with a 440 by 580 minimum. The top selected-message strip is 76 pixels high, followed by a compact 38-pixel toolbar. The transcript consumes all remaining flexible height. The 118-pixel composer band, 64-pixel draft-action area, and 54-pixel status band remain anchored at the bottom.

Horizontal content padding is 18 pixels in the main work areas. The toolbar uses 12 pixels, while the modal settings form uses 24 pixels. Vertical rhythm is compact, generally 3 to 10 pixels between related controls. The transcript stays visually open and scrolls vertically instead of becoming a stack of cards.

The composer is a two-column grid: a fluid multiline text field and a fixed 104-pixel action column. The send button fills the available composer height. Draft actions align to the right beneath a full-width disclosure. Long message subjects, sender metadata, and status text ellipsize rather than breaking the overall frame.

The settings window is a centered modal, 520 by 365 client pixels with a 480 by 390 minimum. Endpoint, model, and API-key fields stack vertically. Data-use guidance and validation errors sit before the bottom-right Save and Cancel actions.

### Named Rules

**The Ordered Boundary Rule.** Keep message identity, conversation, composer, draft controls, and status in that order. This sequence is the interface's security explanation.

**The Transcript-Breathes Rule.** Fixed utility bands yield height to the transcript. Never shrink the conversation into a decorative card or make drafting controls compete with it.

## Elevation & Depth

The application defines no shadows, gradients, blur, overlays, or custom elevation. Depth comes from native window chrome, alternating white and muted system surfaces, and one-pixel control borders. Any outer window shadow is owned by Windows and must not be replicated inside the client area.

### Named Rules

**The Flat-Utility Rule.** App content is flat by default. Use tonal bands and native borders to separate regions, not card shadows or floating panels.

## Shapes

Controls use square native geometry. Text fields have fixed single borders; flat buttons have either a single outline or no border for link-like toolbar actions. Surfaces are rectangular bands with no clipping, pills, avatars, speech bubbles, or decorative silhouettes.

### Named Rules

**The Native Rectangle Rule.** Keep controls square and familiar. Rounded chat bubbles, pill buttons, and oversized AI ornaments contradict the product's utility character.

## Components

### Selected Message Strip

- **Character:** Quiet context anchor, not a card.
- **Structure:** Muted full-width band with a bold, ellipsized subject and a secondary sender line.
- **State:** When no email is available, the same region says "No email selected" and tells the user to select or open one before refreshing.

### Toolbar Links

- **Shape:** Borderless, square, 28-pixel-high buttons.
- **Color:** White canvas with action-blue text in standard contrast.
- **State:** Refresh email and Settings disable while a request is active. Native focus and disabled rendering remain visible.

### Transcript

- **Character:** A spacious, read-only plain-text document.
- **Structure:** Borderless white surface with vertical scrolling and no automatic URL detection.
- **Turns:** Speaker names are bold. "You" uses the action color; "Assistant" uses primary text. System messages are italic secondary text.
- **Accessibility:** Accessible name is "AI chat conversation"; the description identifies it as a plain-text conversation about the selected Outlook email.

### Composer

- **Style:** Multiline square field with a fixed single border, vertical scrolling, and bounded input length.
- **Instruction:** A persistent hint says the field is for questions or draft-text requests and that Ctrl+Enter sends.
- **Focus:** Keep the native Windows focus indication. Do not replace it with color-only styling.
- **Busy State:** Disable the field while waiting, change "Send to AI" to "Cancel," and restore the user's prompt if the request fails, times out, or is discarded.

### Primary Button

- **Shape:** Square, flat, filled action control.
- **Primary:** Action-blue background, system highlight text, bold label.
- **High Contrast:** Replace the fixed fill with the system highlight role.
- **State:** "Send to AI" starts the bounded request. "Cancel" is the only alternate label and cancels the in-flight request.

### Draft Buttons

- **Shape:** Square, flat, 34-pixel-high secondary controls with one-pixel system borders.
- **Labels:** "Open new draft" and "Open reply draft" explicitly describe the result.
- **State:** Both remain disabled until a complete assistant response exists. Reply draft also requires a reply-capable selected message. Both disable during an active request.
- **Boundary:** A click may create, save, and display an unsent Outlook draft. No send, schedule, move, delete, mark, categorize, or source-message modification action belongs in this component family.

### Draft Disclosure

- **Character:** Persistent, quiet, and unambiguous.
- **Copy:** "Drafts use the entire latest assistant response."
- **Placement:** Immediately above the draft buttons so users see the content boundary before acting.

### Status Band

- **Character:** A full-width operational ledger at the bottom of the window.
- **Content:** States selected-message scope, waiting, cancellation, timeout, discarded response, configuration, and unsent-draft outcomes in plain language.
- **Accessibility:** Exposes a status-bar role. Errors use error color plus explicit recovery copy.

### Settings Fields and Actions

- **Fields:** Endpoint, Model, and API key are stacked square inputs with bold labels and accessible descriptions. The API key uses the system password character.
- **Disclosure:** Explain that the selected email and recent conversation go to the configured endpoint, HTTP is loopback-only, and the key is encrypted for the current Windows user.
- **Actions:** Save is primary; Cancel is secondary. Enter activates Save and Escape activates Cancel.
- **Errors:** Validation failures appear inline as an accessible alert without closing the modal.

## Do's and Don'ts

### Do:

- **Do** keep the selected-message identity visible before conversation content.
- **Do** state that only the selected email is available to chat.
- **Do** keep draft controls secondary, explicit, and disabled until a valid assistant response exists.
- **Do** disclose that drafting uses the entire latest assistant response.
- **Do** say "unsent draft" and "for your review" in successful draft status text.
- **Do** inherit Windows system fonts, focus behavior, text scaling, and high-contrast colors.
- **Do** preserve keyboard operation, including Ctrl+Enter to send, Enter to save settings, and Escape to cancel settings.
- **Do** restore user input after request failure, timeout, cancellation, or changed-message discard when applicable.

### Don't:

- **Don't** add a Send Mail control, auto-send behavior, scheduling, source-message mutation, or language that implies any of those capabilities.
- **Don't** expose model tool calls, commands, HTML, clickable output, or rich interactive response cards.
- **Don't** hide which email is loaded or allow a response for a changed email to enter the active conversation.
- **Don't** rely on fixed colors when Windows high contrast is active.
- **Don't** replace native focus and disabled states with color-only signals.
- **Don't** use chat bubbles, avatars, glowing AI motifs, gradients, rounded cards, or ornamental motion.
