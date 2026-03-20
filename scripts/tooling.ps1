Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-HybridRoot {
    return [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
}

function Get-ParentMacroRoot {
    return [System.IO.Path]::GetFullPath((Join-Path (Get-HybridRoot) '..'))
}

function Get-SourceWorkbookPath {
    $parentRoot = Get-ParentMacroRoot
    $candidate = Get-ChildItem -LiteralPath $parentRoot -Filter '*PHI.xlsm' -File | Select-Object -First 1
    if (-not $candidate) {
        throw "Das Quell-Workbook konnte im Ordner '$parentRoot' nicht gefunden werden."
    }
    return $candidate.FullName
}

function Use-PortableNode {
    $root = Get-HybridRoot
    $nodeRoot = Join-Path $root 'tools\node-v24.14.0-win-x64'
    $npmBin = Join-Path $nodeRoot 'node_modules\npm\bin'

    if (-not (Test-Path -LiteralPath (Join-Path $nodeRoot 'node.exe'))) {
        throw "Portable Node wurde nicht gefunden: $nodeRoot"
    }

    $env:PATH = "$nodeRoot;$npmBin;$env:PATH"
    return $nodeRoot
}

function Use-RustGnuToolchain {
    $root = Get-HybridRoot
    $llvmBin = Join-Path $root 'tools\llvm-mingw-20260311-msvcrt-x86_64\bin'
    $selfContained = Join-Path $env:USERPROFILE '.rustup\toolchains\stable-x86_64-pc-windows-gnu\lib\rustlib\x86_64-pc-windows-gnu\lib\self-contained'

    if (-not (Test-Path -LiteralPath (Join-Path $llvmBin 'x86_64-w64-mingw32-gcc.exe'))) {
        throw "Portable LLVM-MinGW Toolchain wurde nicht gefunden: $llvmBin"
    }

    if (-not (Test-Path -LiteralPath $selfContained)) {
        throw "Rust GNU self-contained libs wurden nicht gefunden: $selfContained"
    }

    $env:CARGO_HTTP_CHECK_REVOKE = 'false'
    $env:PATH = "$llvmBin;$env:PATH"
    $env:LIBRARY_PATH = $selfContained
    return $llvmBin
}

function Publish-Frontend {
    $root = Get-HybridRoot
    $frontend = Join-Path $root 'frontend'
    $dist = Join-Path $frontend 'dist'
    $runtimeWeb = Join-Path $root 'runtime\web'

    Use-PortableNode | Out-Null
    Push-Location $frontend
    try {
        & npm.cmd run build | Out-Host
    } finally {
        Pop-Location
    }

    if (Test-Path -LiteralPath $runtimeWeb) {
        Remove-Item -LiteralPath $runtimeWeb -Recurse -Force
    }
    [System.IO.Directory]::CreateDirectory($runtimeWeb) | Out-Null
    Copy-Item -Path (Join-Path $dist '*') -Destination $runtimeWeb -Recurse -Force
    return $runtimeWeb
}

function Build-Backend {
    param([switch]$Release)

    $root = Get-HybridRoot
    $backend = Join-Path $root 'backend'
    Use-RustGnuToolchain | Out-Null

    Push-Location $backend
    try {
        if ($Release) {
            & cargo +stable-x86_64-pc-windows-gnu build --release | Out-Host
            return (Join-Path $backend 'target\release\hybrid_approach.exe')
        }

        & cargo +stable-x86_64-pc-windows-gnu build | Out-Host
        return (Join-Path $backend 'target\debug\hybrid_approach.exe')
    } finally {
        Pop-Location
    }
}
