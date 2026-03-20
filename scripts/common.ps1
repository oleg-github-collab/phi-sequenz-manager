Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RequiredBasFiles = @(
    'Modul1_Konstanten.bas',
    'Modul2_Logger.bas',
    'Modul3_Hilfsfunktionen.bas',
    'Modul4_ImportCSV.bas',
    'Modul5_SequenzSchreiben.bas',
    'Modul6_Hauptsteuerung.bas'
)

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)
}

function New-Utf8Encoding {
    return [System.Text.UTF8Encoding]::new($false)
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Value
    )

    $directory = Split-Path -Parent $Path
    if ($directory) {
        [System.IO.Directory]::CreateDirectory($directory) | Out-Null
    }

    $json = $Value | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($Path, $json, (New-Utf8Encoding))
}

function Read-JsonFileOrNull {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $content = [System.IO.File]::ReadAllText($Path, (New-Utf8Encoding))
    if ([string]::IsNullOrWhiteSpace($content)) {
        return $null
    }

    return $content | ConvertFrom-Json
}

function Get-StableFileHash {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Get-BasModuleHashMap {
    param([Parameter(Mandatory)][string]$BasFolder)

    $hashMap = [ordered]@{}
    foreach ($fileName in $script:RequiredBasFiles) {
        $fullPath = Join-Path $BasFolder $fileName
        if (-not (Test-Path -LiteralPath $fullPath)) {
            throw "BAS-Datei fehlt: $fullPath"
        }
        $hashMap[$fileName] = Get-StableFileHash -Path $fullPath
    }

    return $hashMap
}

function New-HybridTempDirectory {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ("hybrid-approach-" + [guid]::NewGuid().ToString('N'))
    [System.IO.Directory]::CreateDirectory($path) | Out-Null
    return $path
}

function Get-AutomationModuleContent {
    return @'
Attribute VB_Name = "Modul7_Automation"
Option Explicit

Public gAutomationMode As Boolean

Public Sub AutomationSetMode(ByVal enabled As Boolean)
    gAutomationMode = enabled
End Sub

Public Function AutomationMessage(ByVal Prompt As String, _
                                  Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal Title As String = "", _
                                  Optional ByVal HelpFile As String = "", _
                                  Optional ByVal Context As Long = 0) As VbMsgBoxResult
    If gAutomationMode Then
        On Error Resume Next
        If Len(Title) > 0 Then
            Call Log(LVL_INFO, "Automation", "Dialog unterdrückt", Title & " | " & Prompt)
        Else
            Call Log(LVL_INFO, "Automation", "Dialog unterdrückt", Prompt)
        End If
        On Error GoTo 0
        AutomationMessage = vbOK
    Else
        AutomationMessage = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    End If
End Function
'@
}

function New-PatchedModuleSet {
    param(
        [Parameter(Mandatory)][string]$BasFolder,
        [Parameter(Mandatory)][string]$DatPath,
        [Parameter(Mandatory)][string]$TargetFolder
    )

    [System.IO.Directory]::CreateDirectory($TargetFolder) | Out-Null
    $patchedModules = New-Object System.Collections.Generic.List[string]

    foreach ($fileName in $script:RequiredBasFiles) {
        $sourcePath = Join-Path $BasFolder $fileName
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            throw "BAS-Datei fehlt: $sourcePath"
        }

        $content = [System.IO.File]::ReadAllText($sourcePath, (New-Utf8Encoding))

        if ($fileName -eq 'Modul1_Konstanten.bas') {
            $content = [System.Text.RegularExpressions.Regex]::Replace(
                $content,
                '(?m)^Public Const PFAD_DAT\s+As String = ".*"$',
                ('Public Const PFAD_DAT           As String = "{0}"' -f $DatPath)
            )
        }

        if ($fileName -in @('Modul4_ImportCSV.bas', 'Modul5_SequenzSchreiben.bas', 'Modul6_Hauptsteuerung.bas')) {
            $content = $content -replace '\bMsgBox\b', 'AutomationMessage'
        }

        $targetPath = Join-Path $TargetFolder $fileName
        [System.IO.File]::WriteAllText($targetPath, $content, (New-Utf8Encoding))
        $patchedModules.Add($targetPath)
    }

    $automationModulePath = Join-Path $TargetFolder 'Modul7_Automation.bas'
    [System.IO.File]::WriteAllText($automationModulePath, (Get-AutomationModuleContent), (New-Utf8Encoding))
    $patchedModules.Add($automationModulePath)

    return $patchedModules.ToArray()
}

function Remove-ExistingStandardModule {
    param(
        [Parameter(Mandatory)]$VBProject,
        [Parameter(Mandatory)][string]$ModuleName
    )

    foreach ($component in @($VBProject.VBComponents)) {
        if ($component.Name -eq $ModuleName -and [int]$component.Type -eq 1) {
            $VBProject.VBComponents.Remove($component)
            break
        }
    }
}

function New-ExcelApplication {
    param([switch]$Visible)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = [bool]$Visible
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1
    return $excel
}

function Close-ExcelSession {
    param(
        [Parameter()]$Workbook,
        [Parameter()]$Excel
    )

    if ($Workbook) {
        try { $Workbook.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    }

    if ($Excel) {
        try { $Excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

function Import-PatchedModules {
    param(
        [Parameter(Mandatory)]$Workbook,
        [Parameter(Mandatory)][string[]]$ModulePaths
    )

    try {
        $vbProject = $Workbook.VBProject
    } catch {
        throw "Der Zugriff auf das VBA-Projekt ist blockiert. Aktivieren Sie in Excel 'Trust access to the VBA project object model'."
    }

    foreach ($modulePath in $ModulePaths) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($modulePath)
        Remove-ExistingStandardModule -VBProject $vbProject -ModuleName $moduleName
        [void]$vbProject.VBComponents.Import($modulePath)
    }
}

function Compile-VbaProject {
    param(
        [Parameter(Mandatory)]$Excel,
        [Parameter(Mandatory)]$Workbook
    )

    $Workbook.Activate() | Out-Null
    $compileControl = $Excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compileControl) {
        throw 'Der VBA-Compile-Befehl konnte nicht gefunden werden.'
    }

    if ($compileControl.Enabled) {
        [void]$compileControl.Execute()
        Start-Sleep -Milliseconds 1200
    }

    if ($compileControl.Enabled) {
        throw 'Das VBA-Projekt meldet nach dem Compile weiterhin Fehler.'
    }
}

function Get-WorkbookManifestPath {
    param([Parameter(Mandatory)][string]$GeneratedWorkbook)
    return [System.IO.Path]::ChangeExtension($GeneratedWorkbook, '.manifest.json')
}

function Test-WorkbookBuildRequired {
    param(
        [Parameter(Mandatory)][string]$GeneratedWorkbook,
        [Parameter(Mandatory)][string]$ManifestPath,
        [Parameter(Mandatory)][string]$SourceWorkbook,
        [Parameter(Mandatory)][string]$BasFolder,
        [Parameter(Mandatory)][string]$DatPath
    )

    if (-not (Test-Path -LiteralPath $GeneratedWorkbook)) {
        return $true
    }

    $manifest = Read-JsonFileOrNull -Path $ManifestPath
    if ($null -eq $manifest) {
        return $true
    }

    if ($manifest.source_workbook -ne $SourceWorkbook) {
        return $true
    }

    if ($manifest.dat_path -ne $DatPath) {
        return $true
    }

    if ($manifest.source_workbook_hash -ne (Get-StableFileHash -Path $SourceWorkbook)) {
        return $true
    }

    $expectedHashes = Get-BasModuleHashMap -BasFolder $BasFolder
    foreach ($fileName in $script:RequiredBasFiles) {
        if ($manifest.bas_hashes.$fileName -ne $expectedHashes[$fileName]) {
            return $true
        }
    }

    return $false
}

function Invoke-WorkbookBuild {
    param(
        [Parameter(Mandatory)][string]$SourceWorkbook,
        [Parameter(Mandatory)][string]$BasFolder,
        [Parameter(Mandatory)][string]$GeneratedWorkbook,
        [Parameter(Mandatory)][string]$DatPath,
        [Parameter(Mandatory)][string]$ManifestPath,
        [switch]$Visible
    )

    if (-not (Test-Path -LiteralPath $SourceWorkbook)) {
        throw "Quell-Workbook nicht gefunden: $SourceWorkbook"
    }

    $generatedFolder = Split-Path -Parent $GeneratedWorkbook
    if ($generatedFolder) {
        [System.IO.Directory]::CreateDirectory($generatedFolder) | Out-Null
    }

    $tempFolder = New-HybridTempDirectory
    $excel = $null
    $workbook = $null

    try {
        Write-Step "Erzeuge gepatchte BAS-Module."
        $patchedModules = New-PatchedModuleSet -BasFolder $BasFolder -DatPath $DatPath -TargetFolder $tempFolder

        Write-Step "Kopiere Basis-Workbook nach '$GeneratedWorkbook'."
        Copy-Item -LiteralPath $SourceWorkbook -Destination $GeneratedWorkbook -Force

        Write-Step 'Starte Excel für den Build.'
        $excel = New-ExcelApplication -Visible:$Visible
        $workbook = $excel.Workbooks.Open($GeneratedWorkbook, $null, $false)

        Write-Step 'Importiere gepatchte VBA-Module.'
        Import-PatchedModules -Workbook $workbook -ModulePaths $patchedModules

        Write-Step 'Kompiliere VBA-Projekt.'
        Compile-VbaProject -Excel $excel -Workbook $workbook

        [void]$workbook.Save()
        $workbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        $workbook = $null

        $manifest = [ordered]@{
            built_at = (Get-Date).ToString('o')
            generated_workbook = $GeneratedWorkbook
            source_workbook = $SourceWorkbook
            source_workbook_hash = Get-StableFileHash -Path $SourceWorkbook
            dat_path = $DatPath
            bas_hashes = Get-BasModuleHashMap -BasFolder $BasFolder
        }
        Write-JsonFile -Path $ManifestPath -Value $manifest

        return $manifest
    } finally {
        Close-ExcelSession -Workbook $workbook -Excel $excel
        if (Test-Path -LiteralPath $tempFolder) {
            Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-WorksheetOrNull {
    param(
        [Parameter(Mandatory)]$Workbook,
        [Parameter(Mandatory)][string]$WorksheetName
    )

    try {
        return $Workbook.Worksheets.Item($WorksheetName)
    } catch {
        return $null
    }
}

function Convert-RangeToRows {
    param([Parameter(Mandatory)]$Range)

    $value = $Range.Value2
    if ($null -eq $value) {
        return @()
    }

    if (-not ($value -is [System.Array])) {
        return ,@([string]$value)
    }

    if ($value.Rank -eq 1) {
        $row = @()
        foreach ($cell in $value) {
            $row += [string]$(if ($null -eq $cell) { '' } else { $cell })
        }
        return ,$row
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $rowLower = $value.GetLowerBound(0)
    $rowUpper = $value.GetUpperBound(0)
    $colLower = $value.GetLowerBound(1)
    $colUpper = $value.GetUpperBound(1)

    for ($rowIndex = $rowLower; $rowIndex -le $rowUpper; $rowIndex++) {
        $row = New-Object System.Collections.Generic.List[string]
        for ($colIndex = $colLower; $colIndex -le $colUpper; $colIndex++) {
            $cell = $value.GetValue($rowIndex, $colIndex)
            if ($null -eq $cell) {
                $row.Add('')
            } else {
                $row.Add([string]$cell)
            }
        }
        $rows.Add($row.ToArray())
    }

    return $rows.ToArray()
}

function Convert-ToCsvCell {
    param([AllowNull()][string]$Value)

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ($text.Contains('"')) {
        $text = $text.Replace('"', '""')
    }

    if ($text.Contains(';') -or $text.Contains("`n") -or $text.Contains("`r") -or $text.Contains('"')) {
        return '"' + $text + '"'
    }

    return $text
}

function Export-WorksheetCsv {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][string]$TargetPath
    )

    $rows = Convert-RangeToRows -Range $Worksheet.UsedRange
    $directory = Split-Path -Parent $TargetPath
    if ($directory) {
        [System.IO.Directory]::CreateDirectory($directory) | Out-Null
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($row in $rows) {
        $cells = foreach ($cell in $row) {
            Convert-ToCsvCell -Value $cell
        }
        $lines.Add(($cells -join ';'))
    }

    [System.IO.File]::WriteAllLines($TargetPath, $lines, (New-Utf8Encoding))
}

function Get-WorksheetPreview {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [int]$MaxRows = 25
    )

    $rows = Convert-RangeToRows -Range $Worksheet.UsedRange
    if ($rows.Count -eq 0) {
        return @{
            sheet_name = $Worksheet.Name
            headers = @()
            rows = @()
        }
    }

    $headers = @()
    foreach ($value in $rows[0]) {
        $headers += [string]$value
    }

    $previewRows = New-Object System.Collections.Generic.List[object]
    $upper = [Math]::Min($rows.Count - 1, $MaxRows)
    for ($index = 1; $index -le $upper; $index++) {
        $previewRows.Add($rows[$index])
    }

    return @{
        sheet_name = $Worksheet.Name
        headers = $headers
        rows = $previewRows.ToArray()
    }
}

function Get-LogOutcome {
    param([Parameter(Mandatory)]$Worksheet)

    $rows = Convert-RangeToRows -Range $Worksheet.UsedRange
    if ($rows.Count -lt 2) {
        return @{
            success = $false
            message = 'Log-Sheet ist leer.'
        }
    }

    $headers = $rows[0]
    $levelIndex = [Array]::IndexOf($headers, 'Level')
    $moduleIndex = [Array]::IndexOf($headers, 'Modul')
    $messageIndex = [Array]::IndexOf($headers, 'Nachricht')

    if ($levelIndex -lt 0 -or $moduleIndex -lt 0 -or $messageIndex -lt 0) {
        return @{
            success = $false
            message = 'Log-Sheet hat nicht das erwartete Format.'
        }
    }

    for ($index = $rows.Count - 1; $index -ge 1; $index--) {
        $row = $rows[$index]
        if ($row[$moduleIndex] -ne 'Logger') {
            continue
        }

        $logMessage = [string]$row[$messageIndex]
        $level = [string]$row[$levelIndex]
        if ($logMessage -like 'Session * beendet') {
            return @{
                success = ($level -eq 'OK')
                message = $logMessage
            }
        }
    }

    return @{
        success = $false
        message = 'Im Log wurde kein Session-Abschluss gefunden.'
    }
}

function Invoke-ExcelMacroPipeline {
    param(
        [Parameter(Mandatory)][string]$GeneratedWorkbook,
        [Parameter(Mandatory)][string]$MacroName,
        [Parameter(Mandatory)][string]$OutputDir,
        [switch]$Visible
    )

    $excel = $null
    $workbook = $null

    try {
        [System.IO.Directory]::CreateDirectory($OutputDir) | Out-Null
        Write-Step 'Starte Excel für den Makrolauf.'
        $excel = New-ExcelApplication -Visible:$Visible
        $workbook = $excel.Workbooks.Open($GeneratedWorkbook, $null, $false)

        try {
            [void]$excel.Run(("'" + $workbook.Name + "'!AutomationSetMode"), $true)
        } catch {
            throw "Das Automationsmodul konnte nicht initialisiert werden: $($_.Exception.Message)"
        }

        Write-Step ("Führe Makro '{0}' aus." -f $MacroName)
        [void]$excel.Run(("'" + $workbook.Name + "'!" + $MacroName))
        [void]$workbook.Save()

        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $artifactRoot = Join-Path $OutputDir ("run-" + $timestamp)
        [System.IO.Directory]::CreateDirectory($artifactRoot) | Out-Null

        $artifacts = New-Object System.Collections.Generic.List[object]
        $artifacts.Add(@{ label = 'Workbook'; path = $GeneratedWorkbook })

        foreach ($sheetName in @('Needle1', 'Daten', '_Execution_Log')) {
            $sheet = Get-WorksheetOrNull -Workbook $workbook -WorksheetName $sheetName
            if ($null -ne $sheet) {
                $targetPath = Join-Path $artifactRoot ($sheetName + '.csv')
                Export-WorksheetCsv -Worksheet $sheet -TargetPath $targetPath
                $artifacts.Add(@{ label = $sheetName; path = $targetPath })
            }
        }

        $logSheet = Get-WorksheetOrNull -Workbook $workbook -WorksheetName '_Execution_Log'
        if ($null -eq $logSheet) {
            throw "Das Log-Sheet '_Execution_Log' wurde nicht gefunden."
        }

        $logOutcome = Get-LogOutcome -Worksheet $logSheet
        $needleSheet = Get-WorksheetOrNull -Workbook $workbook -WorksheetName 'Needle1'
        $preview = if ($null -ne $needleSheet) {
            Get-WorksheetPreview -Worksheet $needleSheet -MaxRows 30
        } else {
            @{
                sheet_name = 'Needle1'
                headers = @()
                rows = @()
            }
        }

        return @{
            success = [bool]$logOutcome.success
            status_message = [string]$logOutcome.message
            artifacts = $artifacts.ToArray()
            preview = $preview
        }
    } finally {
        Close-ExcelSession -Workbook $workbook -Excel $excel
    }
}
