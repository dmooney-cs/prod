# CS-Toolbox-Launcher-ZeroTouch-GitHubHash.ps1
# - Silent by default (no output/no progress)
# - Downloads toolbox.zip
# - Pulls expected SHA-256 from GitHub RAW
# - Compares to actual file hash
# - Optional: show hashes and/or require confirmation
# - Extracts + normalizes + unblocks
# - Dot-sources launcher in SAME PowerShell session

param(
    [switch]$SkipHashCheck = $false,
    [switch]$ShowHashes = $false,
    [switch]$ConfirmHashes = $false
)

#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'  # no download progress

# --------------------------
# Config
# --------------------------
$ZipUrl       = 'https://github.com/dmooney-cs/prod/raw/refs/heads/main/prod-01-01.zip'
$ZipPath      = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath  = 'C:\CS-Toolbox-TEMP'
$DestRoot     = Join-Path $ExtractPath 'prod-01-01'
$Launcher     = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

# GitHub RAW URL to expected SHA-256 (plain text)
# Example:
# https://raw.githubusercontent.com/YourOrg/CS-Toolbox/main/hash/prod-01-01.sha256
$HashUrl      = 'https://raw.githubusercontent.com/YourOrg/CS-Toolbox/main/hash/prod-01-01.sha256'

# --------------------------
# Output control
# --------------------------
$script:AllowOutput = [bool]($ShowHashes -or $ConfirmHashes)
function Out-Maybe {
    param([string]$Text)
    if ($script:AllowOutput) { Write-Host $Text }
}

# --------------------------
# Helpers (quiet networking)
# --------------------------
function Invoke-DownloadQuiet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$Attempts = 3
    )

    Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }

            if ((Test-Path -LiteralPath $OutFile) -and ((Get-Item -LiteralPath $OutFile).Length -gt 0)) {
                return $true
            }
        } catch {
            Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }
    }
    return $false
}

function Get-RemoteSha256Quiet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$Attempts = 3
    )

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }

            $txt = ($r.Content | Out-String)
            if ([string]::IsNullOrWhiteSpace($txt)) { throw "empty" }

            # First non-empty line, accept either:
            # HASH
            # HASH  filename
            $line = ($txt -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1).Trim()
            $hash = ($line -split '\s+')[0].Trim().ToUpper()

            if ($hash -match '^[0-9A-F]{64}$') { return $hash }
            throw "invalid sha256 format"
        } catch {
            Start-Sleep -Seconds 1
        }
    }

    return $null
}

function Move-ContentsSilent {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Target
    )
    if (-not (Test-Path -LiteralPath $Source)) { return }
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        try { Move-Item -LiteralPath $_.FullName -Destination $Target -Force -ErrorAction Stop } catch { }
    }
}

# --------------------------
# Execution (silent unless switches enable output)
# --------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

try { New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null } catch { return }
try { Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

# Download ZIP
if (-not (Invoke-DownloadQuiet -Uri $ZipUrl -OutFile $ZipPath -Attempts 3)) { return }

$expected = $null
$actual   = $null

if (-not $SkipHashCheck) {
    $expected = Get-RemoteSha256Quiet -Uri $HashUrl -Attempts 3
    if ([string]::IsNullOrWhiteSpace($expected)) {
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }

    try {
        $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
    } catch {
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }

    if ($ShowHashes -or $ConfirmHashes) {
        Out-Maybe ""
        Out-Maybe "Expected (GitHub): $expected"
        Out-Maybe "Actual   (ZIP)   : $actual"
        Out-Maybe ""
    }

    if ($actual -ne $expected) {
        # If you're showing hashes, you’ll already see the mismatch; otherwise remain silent.
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }

    if ($ConfirmHashes) {
        $resp = Read-Host "Hashes match. Proceed with extract + launch? (Y/N)"
        if ($resp -notin @('Y','y')) {
            return
        }
    }
} else {
    if ($ShowHashes -or $ConfirmHashes) {
        Out-Maybe ""
        Out-Maybe "Hash check skipped (-SkipHashCheck)."
        Out-Maybe ""
    }
    if ($ConfirmHashes) {
        $resp = Read-Host "Proceed without hash verification? (Y/N)"
        if ($resp -notin @('Y','y')) { return }
    }
}

# Extract
try { Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force } catch { return }

# Ensure destination
try { New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null } catch { return }

# Normalize ZIP structure
if (-not (Test-Path -LiteralPath (Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'))) {
    $dirs  = Get-ChildItem -LiteralPath $ExtractPath -Directory -Force | Where-Object { $_.FullName -ne $DestRoot }
    $files = Get-ChildItem -LiteralPath $ExtractPath -File -Force | Where-Object { $_.FullName -ne $ZipPath }

    if ($dirs.Count -eq 1 -and $files.Count -eq 0) {
        Move-ContentsSilent -Source $dirs[0].FullName -Target $DestRoot
        Remove-Item -LiteralPath $dirs[0].FullName -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        foreach ($d in $dirs) { Move-ContentsSilent -Source $d.FullName -Target $DestRoot }
        foreach ($f in $files) {
            try { Move-Item -LiteralPath $f.FullName -Destination $DestRoot -Force -ErrorAction Stop } catch { }
        }
        foreach ($d in $dirs) {
            Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Unblock (silent best-effort)
try {
    Get-ChildItem -LiteralPath $DestRoot -Recurse -Force -File |
        ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { } }
} catch { }

# Launch (same session)
if (Test-Path -LiteralPath $Launcher) {
    . $Launcher
}
