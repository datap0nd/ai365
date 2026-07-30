$ErrorActionPreference = "Stop"

$encodedPath = Join-Path $PSScriptRoot "..\build\OutlookLocalAIChat.snk.base64"
$keyPath = Join-Path $PSScriptRoot "..\src\OutlookLocalAIChat\Properties\OutlookLocalAIChat.snk"
$encoded = (Get-Content $encodedPath -Raw).Trim()
$bytes = [Convert]::FromBase64String($encoded)
[IO.File]::WriteAllBytes($keyPath, $bytes)
Write-Host "Restored the strong-name key used for stable COM identity."
