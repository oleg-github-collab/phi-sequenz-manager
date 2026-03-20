param(
    [int]$Port = 8765,
    [switch]$OpenBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$candidatePaths = @(
    (Join-Path $root 'hybrid_approach.exe'),
    (Join-Path $root 'backend\target\release\hybrid_approach.exe'),
    (Join-Path $root 'backend\target\debug\hybrid_approach.exe')
)

$binary = $candidatePaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $binary) {
    throw 'Es wurde kein startfähiges Backend-Binary gefunden.'
}

if ($OpenBrowser) {
    Start-Process "http://127.0.0.1:$Port" | Out-Null
}

& $binary serve --root $root --port $Port
