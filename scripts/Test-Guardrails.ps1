$ErrorActionPreference = "Stop"

$sourceRoot = Join-Path $PSScriptRoot "..\src\OutlookLocalAIChat"
$sourceFiles = Get-ChildItem $sourceRoot -Recurse -Filter *.cs

$forbidden = @(
    "\.Send\s*\(",
    "\.Delete\s*\(",
    "\.Move\s*\(",
    "\.Submit\s*\(",
    "olFolderOutbox",
    "SendAndReceive"
)

foreach ($pattern in $forbidden) {
    $matches = $sourceFiles | Select-String -Pattern $pattern |
        Where-Object { $_.Line -notmatch 'File\.Delete' }
    if ($matches) {
        $matches | ForEach-Object {
            Write-Error "Forbidden Outlook capability: $($_.Path):$($_.LineNumber)"
        }
    }
}

$clientPath = Join-Path $sourceRoot "Chat\OpenAiCompatibleClient.cs"
$factoryPath = Join-Path $sourceRoot "Chat\ChatRequestFactory.cs"
$catalogPath = Join-Path $sourceRoot "Chat\MailboxToolCatalog.cs"
$draftCatalogPath = Join-Path $sourceRoot "Chat\DraftToolCatalog.cs"
$toolHostPath = Join-Path $sourceRoot "Outlook\MailboxToolHost.cs"
$mailboxContextPath = Join-Path $sourceRoot "Outlook\MailboxContextService.cs"
$draftToolHostPath = Join-Path $sourceRoot "Outlook\DraftToolHost.cs"
$chatPanePath = Join-Path $sourceRoot "UI\ChatPane.cs"
$intentPath = Join-Path $sourceRoot "Security\DraftIntentPolicy.cs"
$workingSetPath = Join-Path $sourceRoot "Outlook\MailboxWorkingSet.cs"
$safeModelTextPath = Join-Path $sourceRoot "Security\SafeModelText.cs"
$toneFactoryPath = Join-Path $sourceRoot "Chat\ToneProfileRequestFactory.cs"
$externalContextPath = Join-Path $sourceRoot "Chat\ExternalContextDocument.cs"
$settingsWindowPath = Join-Path $sourceRoot "UI\SettingsWindow.cs"
$settingsStorePath = Join-Path $sourceRoot "Configuration\SettingsStore.cs"
$addInPath = Join-Path $sourceRoot "AddIn.cs"
$catalogSource = Get-Content $catalogPath -Raw
$draftCatalogSource = Get-Content $draftCatalogPath -Raw
$modelFacingSource =
    (Get-Content $clientPath -Raw) +
    (Get-Content $factoryPath -Raw) +
    $catalogSource +
    $draftCatalogSource

$toolNames = [regex]::Matches(
    $catalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
$approvedToolNames = @(
    "read_messages",
    "read_thread",
    "search_mailbox"
) | Sort-Object
if (Compare-Object $toolNames $approvedToolNames) {
    throw "Mailbox tool catalog contains an unexpected capability."
}

$draftToolNames = @(
    [regex]::Matches(
        $draftCatalogSource,
        'public const string \w+ = "([^"]+)";'
    ) | ForEach-Object { $_.Groups[1].Value }
)
$approvedDraftToolNames = @(
    "create_draft",
    "update_draft"
) | Sort-Object
if (Compare-Object ($draftToolNames | Sort-Object) $approvedDraftToolNames) {
    throw "Draft tool catalog contains an unexpected capability."
}

foreach ($capability in @(
    "DraftService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($modelFacingSource.Contains($capability)) {
        throw "Model-facing source references forbidden capability $capability."
    }
}

$toolHostSource = Get-Content $toolHostPath -Raw
foreach ($capability in @(
    "DraftService",
    "CreateReplyDraft",
    "CreateNewDraft"
)) {
    if ($toolHostSource.Contains($capability)) {
        throw "Model-invoked mailbox host references draft capability $capability."
    }
}


$draftToolHostSource = Get-Content $draftToolHostPath -Raw
foreach ($requiredBoundary in @(
    "OneShotDraftAuthorization",
    "authorization.TryConsume()",
    "authorization.MarkCreated()",
    "authorization.MarkUpdated()",
    "DRAFT_PERMISSION_NOT_AVAILABLE",
    "DRAFT_UPDATE_NOT_AVAILABLE",
    "DRAFT_ALREADY_LINKED",
    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
    "DRAFT_REPLY_HANDLE_REQUIRED",
    "DRAFT_REPLY_HANDLE_UNKNOWN",
    'GetString(arguments, "reply_handle")'
)) {
    if (-not $draftToolHostSource.Contains($requiredBoundary)) {
        throw "Draft tool host is missing boundary $requiredBoundary."
    }
}

$factorySource = Get-Content $factoryPath -Raw
if (-not $factorySource.Contains("if (allowDraftCreate && activeDraft == null)") -or
    -not $factorySource.Contains("else if (allowDraftUpdate && activeDraft != null)") -or
    -not $factorySource.Contains("DraftToolCatalog.CreateDefinition()") -or
    -not $factorySource.Contains("DraftToolCatalog.UpdateDefinition()")) {
    throw "Draft tool exposure is not conditionally authorized."
}

$chatPaneSource = Get-Content $chatPanePath -Raw
if (-not $toolHostSource.Contains("ResolveHandle") -or
    -not $chatPaneSource.Contains("mailboxTools.ResolveHandle")) {
    throw "Reply drafts are not bound to request-scoped mailbox handles."
}

if ($chatPaneSource.Contains("_allowOneDraft") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsCreate(prompt)") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsUpdate(prompt)") -or
    -not (Test-Path $intentPath) -or
    -not $chatPaneSource.Contains("UpdateDraftState()")) {
    throw "Automatic local draft-intent authorization is incomplete."
}

$workingSetSource = Get-Content $workingSetPath -Raw
$mailboxContextSource = Get-Content $mailboxContextPath -Raw
$workingSetBoundarySource =
    $workingSetSource +
    $toolHostSource
foreach ($requiredBoundary in @(
    "public const int RecommendedMaxMessages = 10",
    "LimitOverrides.WorkingSetMessages",
    "MAILBOX_WORKING_SET_LOCKED",
    "MAILBOX_CONTEXT_LIMIT_REACHED",
    "MAILBOX_SEARCH_LIMIT_REACHED",
    "_loadedBodyHandles",
    "_searchExecuted"
)) {
    if (-not $workingSetBoundarySource.Contains($requiredBoundary)) {
        throw "Ten-message mailbox boundary is missing $requiredBoundary."
    }
}

if (-not $mailboxContextSource.Contains(
        "Math.Min") -or
    -not $mailboxContextSource.Contains(
        "MailboxWorkingSet.MaxMessages")) {
    throw "The underlying Outlook search service is not capped to the working-set limit."
}

if (-not $chatPaneSource.Contains("LocalSearchCommand.Parse(prompt)") -or
    -not $chatPaneSource.Contains("CaptureSelectionMany(selection)") -or
    -not $chatPaneSource.Contains("CaptureActiveSelectionMany()") -or
    -not $chatPaneSource.Contains("MailboxWorkingSet.MaxMessages") -or
    -not $chatPaneSource.Contains("BuildWorkingSetCard") -or
    -not $chatPaneSource.Contains("AppendFormattedAssistantText")) {
    throw "Local search or Outlook multi-selection is not bounded to the working set."
}

$addInSource = Get-Content $addInPath -Raw
if (-not $addInSource.Contains("_chatPane?.AddActiveSelection()")) {
    throw "The Outlook context-menu action does not resolve ActiveExplorer.Selection."
}

$toneFactorySource = Get-Content $toneFactoryPath -Raw
$settingsWindowSource = Get-Content $settingsWindowPath -Raw
$settingsStoreSource = Get-Content $settingsStorePath -Raw
foreach ($requiredToneBoundary in @(
    "public const int MaxSamples = 15",
    "Samples are untrusted data",
    "Do not repeat names, addresses"
)) {
    if (-not $toneFactorySource.Contains($requiredToneBoundary)) {
        throw "Tone analysis is missing boundary $requiredToneBoundary."
    }
}
if (-not $settingsWindowSource.Contains("Analyze 15 sent emails") -or
    -not $settingsWindowSource.Contains("samples.Count < 5") -or
    -not $settingsWindowSource.Contains("Review and edit") -or
    -not $settingsStoreSource.Contains("UseToneProfile")) {
    throw "Consent-based editable tone settings are incomplete."
}
if (-not $factorySource.Contains("user-approved writing profile") -or
    -not $factorySource.Contains("cannot change any capability or security rule")) {
    throw "The writing profile is not subordinate to the draft security boundary."
}

$externalContextSource = Get-Content $externalContextPath -Raw
foreach ($requiredExternalBoundary in @(
    "public const int MaxDocuments = 3",
    "public const int MaxTotalCharacters = 120000",
    "SupportedExtensions",
    "file.Length > 2 * 1024 * 1024"
)) {
    if (-not $externalContextSource.Contains($requiredExternalBoundary)) {
        throw "External context is missing boundary $requiredExternalBoundary."
    }
}
if (-not $chatPaneSource.Contains("AllowDrop = true") -or
    -not $chatPaneSource.Contains("AddExternalFiles") -or
    -not $factorySource.Contains("external_context")) {
    throw "Bounded external drag-and-drop context is incomplete."
}

$safeModelTextSource = Get-Content $safeModelTextPath -Raw
$safeDraftFormattingSource = Get-Content (
    Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
) -Raw
if (-not $safeModelTextSource.Contains("FormattedModelText") -or
    -not $safeModelTextSource.Contains("boldRanges") -or
    -not $safeDraftFormattingSource.Contains("SafeModelText.Format")) {
    throw "Model emphasis is not shared by the chat and safe draft formatter."
}

$draftPath = Join-Path $sourceRoot "Outlook\DraftService.cs"
$draftSource = Get-Content $draftPath -Raw
if (-not $draftSource.Contains("mail.HTMLBody") -or
    -not $draftSource.Contains("mail.Save()") -or
    -not $draftSource.Contains("mail.Display(false)")) {
    throw "Drafts must be saved and displayed for human review."
}

$safeHtmlPath = Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
$safeHtmlSource = Get-Content $safeHtmlPath -Raw
if (-not $safeHtmlSource.Contains("WebUtility.HtmlEncode") -or
    -not $safeHtmlSource.Contains('output.Append("<strong>")') -or
    -not $safeHtmlSource.Contains('"<h2 style=') -or
    -not $safeHtmlSource.Contains('"<ul style=') -or
    -not $safeHtmlSource.Contains('"<hr style=') -or
    -not $safeHtmlSource.Contains('"<table style=') -or
    $draftCatalogSource.Contains('"html"')) {
    throw "Draft formatting must remain locally encoded and structurally bounded."
}

# ---- AI365 suite guardrails: Excel/PowerPoint hosts and MCP ----

# The document-side hosts may never save, print, protect, close, or
# quit the user's files or applications. (Outlook draft Save stays
# allowed in the Outlook folder, where it persists the reviewed
# unsent draft.)
$officeForbidden = @(
    "\.Save\s*\(",
    "\.SaveAs",
    "\.SaveCopyAs",
    "\.Quit\s*\(",
    "PrintOut",
    "SendMail",
    "\.Protect",
    "\.Unprotect",
    "\.Close\s*\("
)
$officeGuardedFiles = @(
    (Join-Path $sourceRoot "Office\WorkbookToolHost.cs"),
    (Join-Path $sourceRoot "Office\WorkbookDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\PresentationToolHost.cs"),
    (Join-Path $sourceRoot "Office\PresentationDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\WordToolHost.cs"),
    (Join-Path $sourceRoot "Office\WordDraftWriter.cs"),
    (Join-Path $sourceRoot "Office\DraftTextLayout.cs"),
    (Join-Path $sourceRoot "Office\DraftChartTypes.cs"),
    (Join-Path $sourceRoot "Office\DocumentDraftHost.cs"),
    (Join-Path $sourceRoot "ExcelAddIn.cs"),
    (Join-Path $sourceRoot "PowerPointAddIn.cs"),
    (Join-Path $sourceRoot "WordAddIn.cs"),
    (Join-Path $sourceRoot "UI\OfficeChatPane.cs"),
    (Join-Path $sourceRoot "UI\TaskPaneRegistry.cs")
)
foreach ($guardedFile in $officeGuardedFiles) {
    foreach ($pattern in $officeForbidden) {
        # dataWorkbook.Close closes only a chart's own embedded
        # data-grid workbook inside an unsaved draft presentation,
        # never a user file.
        $hits = Select-String -Path $guardedFile -Pattern $pattern |
            Where-Object {
                $_.Line -notmatch '_settingsStore\.Save' -and
                $_.Line -notmatch 'SuiteExchange\.Save' -and
                $_.Line -notmatch 'dataWorkbook\.Close'
            }
        if ($hits) {
            throw "Forbidden document capability $pattern in $guardedFile."
        }
    }
}

$workbookCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\WorkbookToolCatalog.cs") -Raw
$presentationCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\PresentationToolCatalog.cs") -Raw
$crossAppCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\CrossAppToolCatalog.cs") -Raw
$wordCatalogSource = Get-Content (
    Join-Path $sourceRoot "Chat\WordToolCatalog.cs") -Raw
$documentFactorySource = Get-Content (
    Join-Path $sourceRoot "Chat\DocumentChatRequestFactory.cs") -Raw

$workbookToolNames = [regex]::Matches(
    $workbookCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $workbookToolNames (@(
    "list_worksheets",
    "read_cells",
    "write_draft_sheet",
    "write_cells") | Sort-Object)) {
    throw "Workbook tool catalog contains an unexpected capability."
}

$presentationToolNames = [regex]::Matches(
    $presentationCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $presentationToolNames (@(
    "list_slides",
    "read_slide",
    "add_draft_slides") | Sort-Object)) {
    throw "Presentation tool catalog contains an unexpected capability."
}

$crossAppToolNames = [regex]::Matches(
    $crossAppCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $crossAppToolNames (@(
    "create_email_draft",
    "send_to_powerpoint",
    "send_to_excel",
    "send_to_word") | Sort-Object)) {
    throw "Cross-app tool catalog contains an unexpected capability."
}

$wordToolNames = [regex]::Matches(
    $wordCatalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
if (Compare-Object $wordToolNames (@(
    "read_document",
    "write_draft_document") | Sort-Object)) {
    throw "Word tool catalog contains an unexpected capability."
}

$documentDraftHostSource = Get-Content (
    Join-Path $sourceRoot "Office\DocumentDraftHost.cs") -Raw
foreach ($requiredBoundary in @(
    "OneShotDraftAuthorization",
    "authorization.TryConsume()",
    "authorization.MarkCreated()",
    "DRAFT_PERMISSION_NOT_AVAILABLE",
    "DRAFT_TOOL_MUST_BE_EXCLUSIVE"
)) {
    if (-not $documentDraftHostSource.Contains($requiredBoundary)) {
        throw "Document draft host is missing boundary $requiredBoundary."
    }
}

$workbookWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\WorkbookDraftWriter.cs") -Raw
if (-not $workbookWriterSource.Contains(
        'DraftSheetName = "AI365 Draft"')) {
    throw "The Excel write surface must stay pinned to the AI365 Draft sheet."
}
if (-not $workbookWriterSource.Contains(
        "DraftFormulaPolicy.IsAllowedFormula")) {
    throw "Draft formulas must pass the formula safety policy."
}

$formulaPolicySource = Get-Content (
    Join-Path $sourceRoot "Security\DraftFormulaPolicy.cs") -Raw
foreach ($blockedFunction in @(
    '"WEBSERVICE"',
    '"RTD"',
    '"CALL"',
    '"HYPERLINK"'
)) {
    if (-not $formulaPolicySource.Contains($blockedFunction)) {
        throw "The formula policy no longer blocks $blockedFunction."
    }
}

$presentationWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\PresentationDraftWriter.cs") -Raw
if (-not $presentationWriterSource.Contains(
        'DraftMarker = "[AI365 draft]"')) {
    throw "Appended slides must stay marked as AI365 drafts."
}

$wordWriterSource = Get-Content (
    Join-Path $sourceRoot "Office\WordDraftWriter.cs") -Raw
if (-not $wordWriterSource.Contains(
        'DraftMarker = "[AI365 draft]"') -or
    -not $wordWriterSource.Contains("Documents.Add()")) {
    throw "Word drafts must stay marked, new, and unsaved."
}

foreach ($requiredDocumentBoundary in @(
    "can never send email",
    "The local host recognized an explicit draft request",
    "did not recognize an explicit draft, insert, or",
    "untrusted reference data"
)) {
    if (-not $documentFactorySource.Contains($requiredDocumentBoundary)) {
        throw "Document factory is missing boundary $requiredDocumentBoundary."
    }
}

$officePaneSource = Get-Content (
    Join-Path $sourceRoot "UI\OfficeChatPane.cs") -Raw
if (-not $officePaneSource.Contains(
        "DocumentDraftIntentPolicy.AllowsDraft(prompt)")) {
    throw "Document drafts are not gated by the local intent policy."
}

# User-adjustable limits stay clamped and never touch capability
# caps; the panes must push settings limits into the effective
# values.
$textBoundarySource = Get-Content (
    Join-Path $sourceRoot "Security\TextBoundary.cs") -Raw
foreach ($requiredLimitBoundary in @(
    "MaxPromptCharacters = 16000",
    "MaxAssistantCharactersLimit = 48000",
    "MaxToolRoundsLimit = 8",
    "MaxUserMultiplier = 8",
    "MaxWorkingSetMessages = 50",
    "if (useRecommended)"
)) {
    if (-not $textBoundarySource.Contains($requiredLimitBoundary)) {
        throw "Limit overrides are missing clamp $requiredLimitBoundary."
    }
}
if (-not $chatPaneSource.Contains("ApplyLimits()") -or
    -not $officePaneSource.Contains("ApplyLimits()")) {
    throw "Panes do not apply the configured limits."
}

# MCP tools stay namespaced, bounded, and separated from the draft
# and mailbox capability surfaces.
$mcpHostSource = Get-Content (
    Join-Path $sourceRoot "Chat\McpToolHost.cs") -Raw
foreach ($requiredMcpBoundary in @(
    'ToolPrefix = "mcp_"',
    "untrusted_mcp_data",
    "MaxExposedTools = 40"
)) {
    if (-not $mcpHostSource.Contains($requiredMcpBoundary)) {
        throw "MCP host is missing boundary $requiredMcpBoundary."
    }
}
foreach ($mcpForbidden in @(
    "DraftService",
    "DraftToolHost",
    "DocumentDraftHost",
    "MailboxContextService"
)) {
    if ($mcpHostSource.Contains($mcpForbidden)) {
        throw "MCP host references forbidden capability $mcpForbidden."
    }
}

if (-not $settingsWindowSource.Contains(
        "outside this add-in's guardrails")) {
    throw "The MCP settings page is missing its trust notice."
}

# Administrator policy can only remove capabilities: settings load
# forces Gemini off and the gateway refuses to run under policy.
$settingsStoreSource = Get-Content (
    Join-Path $sourceRoot "Configuration\SettingsStore.cs") -Raw
if (-not $settingsStoreSource.Contains(
        "AdminPolicy.GeminiDisabled")) {
    throw "Settings load must honor the Gemini disable policy."
}
$geminiGatewaySource = Get-Content (
    Join-Path $sourceRoot "Chat\GeminiCodeAssistGateway.cs") -Raw
if (-not $geminiGatewaySource.Contains(
        "GEMINI_DISABLED_BY_POLICY")) {
    throw "The Gemini gateway must refuse to run under policy."
}

# The document-side model-facing sources carry the same capability
# hygiene as the mailbox ones.
$documentModelFacingSource =
    $workbookCatalogSource +
    $presentationCatalogSource +
    $crossAppCatalogSource +
    $wordCatalogSource +
    $documentFactorySource
foreach ($capability in @(
    "DraftService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($documentModelFacingSource.Contains($capability)) {
        throw "Document model-facing source references forbidden capability $capability."
    }
}

Write-Host "PASS: static guardrail scan"
