Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'tooling.ps1')

$root = Get-HybridRoot
$parentRoot = Get-ParentMacroRoot
$sourceWorkbook = Get-SourceWorkbookPath
$sampleDat = Join-Path $root 'tests\fixtures\sample_phi.dat'
$runtimeOutput = Join-Path $root 'runtime\output\smoke'
$generatedWorkbook = Join-Path $root 'runtime\generated\phi_hybrid_runtime.xlsm'
$reportPath = Join-Path $runtimeOutput 'smoke-report.json'
$serverStdOut = Join-Path $root 'runtime\logs\smoke-backend.stdout.log'
$serverStdErr = Join-Path $root 'runtime\logs\smoke-backend.stderr.log'
$backendProcess = $null
$port = 8876

try {
    Write-Host '[Smoke] Patche Unicorn-SDK.'
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'patch-unicorn-sdk.ps1')

    Write-Host '[Smoke] Baue Frontend.'
    Publish-Frontend | Out-Null

    Write-Host '[Smoke] Baue Backend.'
    $binary = Build-Backend

    Write-Host '[Smoke] Führe Excel/PowerShell Pipeline gegen Test-DAT aus.'
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'run-excel-pipeline.ps1') `
        -SourceWorkbook $sourceWorkbook `
        -BasFolder $parentRoot `
        -DatPath $sampleDat `
        -OutputDir $runtimeOutput `
        -GeneratedWorkbook $generatedWorkbook `
        -MacroName 'MakrosAusfuehren' `
        -ReportPath $reportPath `
        -ForceRebuild

    Write-Host '[Smoke] Schreibe Test-Settings für API-Prüfung.'
    $settings = @{
        source_workbook_path = $sourceWorkbook
        bas_folder_path = $parentRoot
        dat_file_path = $sampleDat
        output_dir = $runtimeOutput
        generated_workbook_path = $generatedWorkbook
        macro_name = 'MakrosAusfuehren'
        preferred_mode = 'headless'
        auto_rebuild_workbook = $true
        profiles = @(
            @{
                name = 'Smoke'
                source_workbook_path = $sourceWorkbook
                bas_folder_path = $parentRoot
                dat_file_path = $sampleDat
                output_dir = $runtimeOutput
                macro_name = 'MakrosAusfuehren'
                preferred_mode = 'headless'
            }
        )
        watch = @{
            enabled = $false
            watch_folder = (Join-Path $root 'runtime\inbox')
            debounce_ms = 1500
            auto_run_mode = 'headless'
            pattern = '*.dat'
        }
    }
    [System.IO.File]::WriteAllText(
        (Join-Path $root 'runtime\settings.json'),
        ($settings | ConvertTo-Json -Depth 20),
        [System.Text.UTF8Encoding]::new($false)
    )

    Write-Host '[Smoke] Starte Backend für HTTP-Prüfungen.'
    $arguments = "serve --root `"$root`" --port $port"
    $backendProcess = Start-Process `
        -FilePath $binary `
        -ArgumentList $arguments `
        -WorkingDirectory $root `
        -RedirectStandardOutput $serverStdOut `
        -RedirectStandardError $serverStdErr `
        -PassThru

    $health = $null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        if ($backendProcess.HasExited) {
            $stderr = if (Test-Path -LiteralPath $serverStdErr) { Get-Content -Raw $serverStdErr } else { '' }
            $stdout = if (Test-Path -LiteralPath $serverStdOut) { Get-Content -Raw $serverStdOut } else { '' }
            throw "Backend-Prozess ist vorzeitig beendet.`nSTDOUT:`n$stdout`nSTDERR:`n$stderr"
        }

        try {
            $health = Invoke-RestMethod "http://127.0.0.1:$port/api/health"
            break
        } catch {
            Start-Sleep -Seconds 1
        }
    }

    if ($null -eq $health) {
        throw 'Backend-Health-Endpoint wurde nicht rechtzeitig erreichbar.'
    }

    $dashboard = Invoke-RestMethod "http://127.0.0.1:$port/api/dashboard"
    $validation = Invoke-RestMethod "http://127.0.0.1:$port/api/validate"
    $frontend = Invoke-WebRequest "http://127.0.0.1:$port/" -UseBasicParsing

    Write-Host '[Smoke] Health:'
    $health | ConvertTo-Json -Depth 5
    Write-Host '[Smoke] Dashboard-Jobs:'
    $dashboard.data.jobs.Count
    Write-Host '[Smoke] Validierung OK:'
    $validation.data.ok
    Write-Host '[Smoke] Frontend Titel vorhanden:'
    ($frontend.Content -like '*id="root"*')
    Write-Host '[Smoke] Test erfolgreich abgeschlossen.'
} finally {
    if ($backendProcess -and -not $backendProcess.HasExited) {
        Stop-Process -Id $backendProcess.Id -Force
    }
}
