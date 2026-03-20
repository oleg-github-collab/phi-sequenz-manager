param(
    [Parameter(Mandatory)][string]$SourceWorkbook,
    [Parameter(Mandatory)][string]$BasFolder,
    [Parameter(Mandatory)][string]$DatPath,
    [Parameter(Mandatory)][string]$OutputDir,
    [Parameter(Mandatory)][string]$GeneratedWorkbook,
    [Parameter(Mandatory)][string]$MacroName,
    [Parameter(Mandatory)][string]$ReportPath,
    [switch]$Visible,
    [switch]$ForceRebuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

[System.IO.Directory]::CreateDirectory($OutputDir) | Out-Null
$messages = New-Object System.Collections.Generic.List[string]
$artifacts = New-Object System.Collections.Generic.List[object]
$manifestPath = Get-WorkbookManifestPath -GeneratedWorkbook $GeneratedWorkbook
$preview = $null

try {
    $needsBuild = $ForceRebuild -or (Test-WorkbookBuildRequired `
        -GeneratedWorkbook $GeneratedWorkbook `
        -ManifestPath $manifestPath `
        -SourceWorkbook $SourceWorkbook `
        -BasFolder $BasFolder `
        -DatPath $DatPath)

    if ($needsBuild) {
        Write-Step 'Runtime-Workbook wird neu gebaut.'
        $null = Invoke-WorkbookBuild `
            -SourceWorkbook $SourceWorkbook `
            -BasFolder $BasFolder `
            -GeneratedWorkbook $GeneratedWorkbook `
            -DatPath $DatPath `
            -ManifestPath $manifestPath `
            -Visible:$Visible
        $messages.Add('Runtime-Workbook wurde neu gebaut.')
    } else {
        Write-Step 'Vorhandenes Runtime-Workbook wird wiederverwendet.'
        $messages.Add('Vorhandenes Runtime-Workbook wurde wiederverwendet.')
    }

    if (Test-Path -LiteralPath $GeneratedWorkbook) {
        $artifacts.Add(@{ label = 'Workbook'; path = $GeneratedWorkbook })
    }
    if (Test-Path -LiteralPath $manifestPath) {
        $artifacts.Add(@{ label = 'Manifest'; path = $manifestPath })
    }

    $runResult = Invoke-ExcelMacroPipeline `
        -GeneratedWorkbook $GeneratedWorkbook `
        -MacroName $MacroName `
        -OutputDir $OutputDir `
        -Visible:$Visible

    foreach ($message in @($runResult['status_message'])) {
        if (-not [string]::IsNullOrWhiteSpace($message)) {
            $messages.Add([string]$message)
        }
    }

    foreach ($artifact in $runResult['artifacts']) {
        $artifacts.Add($artifact)
    }

    $preview = $runResult['preview']
    $report = @{
        success = [bool]$runResult['success']
        messages = $messages.ToArray()
        artifacts = $artifacts.ToArray()
        preview = $preview
    }
    Write-JsonFile -Path $ReportPath -Value $report

    if (-not $runResult['success']) {
        throw ([string]$runResult['status_message'])
    }

    Write-Step 'Makrolauf erfolgreich abgeschlossen.'
} catch {
    $messages.Add($_.Exception.Message)
    Write-JsonFile -Path $ReportPath -Value @{
        success = $false
        messages = $messages.ToArray()
        artifacts = $artifacts.ToArray()
        preview = $preview
    }
    throw
}
