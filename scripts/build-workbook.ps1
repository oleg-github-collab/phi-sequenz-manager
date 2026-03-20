param(
    [Parameter(Mandatory)][string]$SourceWorkbook,
    [Parameter(Mandatory)][string]$BasFolder,
    [Parameter(Mandatory)][string]$DatPath,
    [Parameter(Mandatory)][string]$GeneratedWorkbook,
    [string]$ReportPath = '',
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$messages = New-Object System.Collections.Generic.List[string]
$artifacts = New-Object System.Collections.Generic.List[object]
$manifestPath = Get-WorkbookManifestPath -GeneratedWorkbook $GeneratedWorkbook

try {
    Write-Step 'Baue Runtime-Workbook.'
    $manifest = Invoke-WorkbookBuild `
        -SourceWorkbook $SourceWorkbook `
        -BasFolder $BasFolder `
        -GeneratedWorkbook $GeneratedWorkbook `
        -DatPath $DatPath `
        -ManifestPath $manifestPath `
        -Visible:$Visible

    $messages.Add('Runtime-Workbook erfolgreich gebaut.')
    $artifacts.Add(@{ label = 'Workbook'; path = $GeneratedWorkbook })
    $artifacts.Add(@{ label = 'Manifest'; path = $manifestPath })

    if ($ReportPath) {
        Write-JsonFile -Path $ReportPath -Value @{
            success = $true
            messages = $messages.ToArray()
            artifacts = $artifacts.ToArray()
            preview = $null
            manifest = $manifest
        }
    }
} catch {
    $messages.Add($_.Exception.Message)
    if ($ReportPath) {
        Write-JsonFile -Path $ReportPath -Value @{
            success = $false
            messages = $messages.ToArray()
            artifacts = $artifacts.ToArray()
            preview = $null
        }
    }
    throw
}
