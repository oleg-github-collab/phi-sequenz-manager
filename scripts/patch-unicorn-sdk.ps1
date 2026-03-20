param(
    [string]$SourceSdk = (Join-Path $PSScriptRoot '..\artifacts\unicornStudio.umd.js'),
    [string]$TargetSdk = (Join-Path $PSScriptRoot '..\frontend\public\vendor\unicornStudio.patched.js')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sourcePath = [System.IO.Path]::GetFullPath($SourceSdk)
$targetPath = [System.IO.Path]::GetFullPath($TargetSdk)

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Quell-SDK nicht gefunden: $sourcePath"
}

$targetFolder = Split-Path -Parent $targetPath
[System.IO.Directory]::CreateDirectory($targetFolder) | Out-Null

$content = [System.IO.File]::ReadAllText($sourcePath, [System.Text.UTF8Encoding]::new($false))
$pattern = 'let m=r\[E\("ZnJlZVBsYW4="\)\];\(m\|\|r\[E\("aW5jbHVkZUxvZ28="\)\]\)&&function\(e,t\)\{.*?\}\(0,n\),'
$patched = [System.Text.RegularExpressions.Regex]::Replace(
    $content,
    $pattern,
    'let m=!1,',
    [System.Text.RegularExpressions.RegexOptions]::Singleline
)

if ($patched -eq $content) {
    throw 'Das Unicorn-SDK konnte nicht gepatcht werden. Das erwartete Branding-Muster wurde nicht gefunden.'
}

[System.IO.File]::WriteAllText($targetPath, $patched, [System.Text.UTF8Encoding]::new($false))
Write-Host "Gepatchtes Unicorn-SDK geschrieben: $targetPath"
