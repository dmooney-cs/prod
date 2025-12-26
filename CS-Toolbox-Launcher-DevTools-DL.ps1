<# =================================================================================================
 CS-Toolbox-Download-Only-ZeroTouch.ps1  (v1.0 - staging extract, verify, exit)

 Purpose:
  - Downloads toolbox ZIP
  - Verifies SHA-256 (unless skipped)
  - Extracts into C:\CS-Toolbox-TEMP\prod-01-01
  - Normalizes ONE wrapper folder if present
  - Overwrites existing files (never creates "(2)")
  - DOES NOT launch anything
  - Exits cleanly

 Switches:
  -SkipHashCheck
  -ShowHashes
  -ConfirmHashes
  -ExportOnly
================================================================================================= #>

#requires -version 5.1
[CmdletBinding()]
param(
    [switch]$SkipHashCheck = $false,
    [switch]$ShowHashes   = $false,
    [switch]$ConfirmHashes = $false,
    [switch]$ExportOnly   = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --------------------------
# Config
# --------------------------
$ZipUrl      = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath     = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath = 'C:\CS-Toolbox-TEMP'
$DestRoot    = Join-Path $ExtractPath 'prod-01-01'

# GitHub RAW hash file
$HashUrl     = 'https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/devtools.sha256'

# --------------------------
# Output control
# --------------------------
$script:AllowOutput = [bool]($ShowHashes -or $ConfirmHashes)

function Say($m) {
    if ($script:AllowOutput) { Write-Host $m }
}

function Fail($stage, $reason) {
    if ($script:AllowOutput) {
        Write-Host ""
        Write-Host "FAILED at: $stage"
        Write-Host "Reason : $reason"
        Write-Host ""
    }
    return $false
}

# --------------------------
# Helpers
# --------------------------
function Invoke-DownloadQuiet {
    param(
        [string]$Uri,
        [string]$OutFile,
        [int]$Attempts = 3
    )

    try { Remove-Item $OutFile -Force -ErrorAction SilentlyContinue } catch {}

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }
            if ((Test-Path $OutFile) -and ((Get-Item $OutFile).Length -gt 0)) {
                return $true
            }
        } catch {
            Start-Sleep 1
        }
    }
    return $false
}

function Get-RemoteSha256Quiet {
    param(
        [string]$Uri,
        [int]$Attempts = 3
    )

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop
            $line = ($r.Content -split "`r?`n" | Where-Object { $_.Trim() })[0]
            $hash = ($line -split '\s+')[0].Trim().ToUpperInvariant()
            if ($hash -match '^[0-9A-F]{64}$') { return $hash }
        } catch {
            Start-Sleep 1
        }
    }
    return $null
}

function New-IsolatedStageFolder {
    param([string]$Parent)
    $path = Join-Path $Parent ("_stage_" + [guid]::NewGuid().ToString("N"))
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

function Resolve-ExtractedRoot {
    param([string]$StageRoot)

    $dirs  = @(Get-ChildItem $StageRoot -Directory -Force)
    $files = @(Get-ChildItem $StageRoot -File -Force)

    if ($dirs.Count -eq 1 -and $files.Count -eq 0) {
        return $dirs[0].FullName
    }
    return $StageRoot
}

function Move-TreeIntoDest {
    param(
        [string]$ContentRoot,
        [string]$DestRoot
    )

    foreach ($item in Get-ChildItem $ContentRoot -Force) {
        try {
            Move-Item $item.FullName $DestRoot -Force
        } catch {
            Copy-Item $item.FullName $DestRoot -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# --------------------------
# Main
# --------------------------
$stage = "init"

try {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $stage = "prepare folders"
    New-Item $ExtractPath -ItemType Directory -Force | Out-Null
    Remove-Item $DestRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item $DestRoot -ItemType Directory -Force | Out-Null

    $stage = "download zip"
    Say "Downloading ZIP..."
    if (-not (Invoke-DownloadQuiet $ZipUrl $ZipPath)) {
        return (Fail $stage "Download failed")
    }

    if (-not $SkipHashCheck) {
        $stage = "hash verification"
        $expected = Get-RemoteSha256Quiet $HashUrl
        if (-not $expected) { return (Fail $stage "Failed to retrieve hash") }

        $actual = (Get-FileHash $ZipPath -Algorithm SHA256).Hash.ToUpperInvariant()

        if ($ShowHashes -or $ConfirmHashes) {
            Write-Host ""
            Write-Host "Expected: $expected"
            Write-Host "Actual  : $actual"
            Write-Host ""
        }

        if ($expected -ne $actual) {
            return (Fail $stage "Hash mismatch")
        }

        if ($ConfirmHashes) {
            if ((Read-Host "Hashes match. Proceed with extract? (Y/N)") -notin @('Y','y')) {
                return $true
            }
        }
    }

    $stage = "extract (staging)"
    $stageRoot = New-IsolatedStageFolder $ExtractPath
    Expand-Archive $ZipPath $stageRoot -Force

    $stage = "resolve root"
    $contentRoot = Resolve-ExtractedRoot $stageRoot

    $stage = "move content"
    Move-TreeIntoDest $contentRoot $DestRoot

    $stage = "cleanup"
    Remove-Item $stageRoot -Recurse -Force -ErrorAction SilentlyContinue

    $stage = "unblock"
    Get-ChildItem $DestRoot -Recurse -File -Force |
        ForEach-Object { try { Unblock-File $_.FullName } catch {} }

    # DONE — NO LAUNCH
    return $true

} catch {
    return (Fail $stage $_.Exception.Message)
}
