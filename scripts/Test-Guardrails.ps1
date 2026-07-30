$ErrorActionPreference = "Stop"

$sourceRoot = Join-Path $PSScriptRoot "..\src\OutlookLocalAIChat"
$sourceFiles = Get-ChildItem $sourceRoot -Recurse -Filter *.cs

$forbidden = @(
    "\.Send\s*\(",
    "\.Delete\s*\(",
    "\.Move\s*\(",
    "olFolderOutbox",
    "SendAndReceive"
)

foreach ($pattern in $forbidden) {
    $matches = $sourceFiles | Select-String -Pattern $pattern
    if ($matches) {
        $matches | ForEach-Object {
            Write-Error "Forbidden Outlook capability: $($_.Path):$($_.LineNumber)"
        }
    }
}

$clientPath = Join-Path $sourceRoot "Chat\OpenAiCompatibleClient.cs"
$factoryPath = Join-Path $sourceRoot "Chat\ChatRequestFactory.cs"
$modelFacingSource = (Get-Content $clientPath -Raw) + (Get-Content $factoryPath -Raw)

foreach ($field in @('"tools"', '"tool_choice"', '"function_call"')) {
    if ($modelFacingSource.Contains($field)) {
        throw "Model-facing source contains forbidden request field $field."
    }
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

$draftPath = Join-Path $sourceRoot "Outlook\DraftService.cs"
$draftSource = Get-Content $draftPath -Raw
if (-not $draftSource.Contains("replyMail.Save()") -or
    -not $draftSource.Contains("replyMail.Display(false)")) {
    throw "Reply draft must be saved and displayed for human review."
}

Write-Host "PASS: static guardrail scan"
