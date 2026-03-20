param(
    [int]$Port = 8765,
    [switch]$Release,
    [switch]$OpenBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'tooling.ps1')

$root = Get-HybridRoot

Write-Host '[Start] Patche lokales Unicorn-SDK.'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'patch-unicorn-sdk.ps1')

Write-Host '[Start] Baue Frontend und kopiere Web-Dateien.'
$runtimeWeb = Publish-Frontend

Write-Host '[Start] Baue Rust-Backend.'
$binary = Build-Backend -Release:$Release

if (-not (Test-Path -LiteralPath $binary)) {
    throw "Backend-Binary wurde nicht gefunden: $binary"
}

$arguments = "serve --root `"$root`" --port $Port"
Write-Host "[Start] Starte Hybrid-Server auf http://127.0.0.1:$Port"
Start-Process -FilePath $binary -ArgumentList $arguments -WorkingDirectory $root | Out-Null

if ($OpenBrowser) {
    Start-Process "http://127.0.0.1:$Port" | Out-Null
}

Write-Host "Frontend ausgeliefert aus: $runtimeWeb"
Write-Host "Serverprozess gestartet. Wenn gewünscht, Browser öffnen: http://127.0.0.1:$Port"
