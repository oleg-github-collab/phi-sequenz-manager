param([switch]$Release)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'tooling.ps1')

$binaryPath = Build-Backend -Release:$Release
Write-Host "Backend gebaut: $binaryPath"
