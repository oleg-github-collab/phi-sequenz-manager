Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'tooling.ps1')

$root = Get-HybridRoot
$releaseRoot = Join-Path $root 'runtime\releases\hybrid-approach-local-win11'
$zipPath = Join-Path $root 'runtime\releases\hybrid-approach-local-win11.zip'

Write-Host '[Release] Patche lokales Unicorn-SDK.'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'patch-unicorn-sdk.ps1')

Write-Host '[Release] Baue Frontend.'
Publish-Frontend | Out-Null

Write-Host '[Release] Baue Backend im Release-Modus.'
$binary = Build-Backend -Release

if (Test-Path -LiteralPath $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
[System.IO.Directory]::CreateDirectory($releaseRoot) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'runtime\web')) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'runtime\output')) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'runtime\generated')) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'runtime\logs')) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'runtime\inbox')) | Out-Null
[System.IO.Directory]::CreateDirectory((Join-Path $releaseRoot 'scripts')) | Out-Null

Copy-Item -LiteralPath $binary -Destination (Join-Path $releaseRoot 'hybrid_approach.exe') -Force
Copy-Item -Path (Join-Path $root 'runtime\web\*') -Destination (Join-Path $releaseRoot 'runtime\web') -Recurse -Force
Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination (Join-Path $releaseRoot 'README.md') -Force

foreach ($scriptName in @(
    'run-release.ps1',
    'build-workbook.ps1',
    'run-excel-pipeline.ps1',
    'common.ps1'
)) {
    Copy-Item -LiteralPath (Join-Path $root "scripts\$scriptName") -Destination (Join-Path $releaseRoot "scripts\$scriptName") -Force
}

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path (Join-Path $releaseRoot '*') -DestinationPath $zipPath

Write-Host "Lokales Release erstellt:"
Write-Host "  Ordner: $releaseRoot"
Write-Host "  Zip:    $zipPath"
