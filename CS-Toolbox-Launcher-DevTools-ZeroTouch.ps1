# CS-Toolbox-Launcher-ZeroTouch-GitHubHash-SILENT.ps1
# Fully silent bootstrapper (NO output, NO prompts, NO progress)
# - Downloads toolbox.zip
# - Fetches expected SHA-256 from GitHub (raw URL)
# - Verifies ZIP hash matches expected
# - Extracts + normalizes
# - Unblocks
# - Dot-sources toolbox launcher in SAME PowerShell session
# - Any failure exits silently

param(
    [switch]$SkipHashCheck = $false
)

#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --------------------------
# Config
# --------------------------
$ZipUrl       = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath      = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath  = 'C:\CS-Toolbox-TEMP'
$DestRoot     = Join-Path $ExtractPath 'prod-01-01'
$Launcher     = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

# GitHub RAW URL to expected SHA-256 (plain text)
# Example:
# https://raw.githubusercontent.com/YourOrg/CS-Toolbox/main/hash/prod-01-01.sha256
$HashUrl      = 'https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/prod-01-01.sha256'

# --------------------------
# Helpers (silent)
# --------------------------
function Invoke-DownloadSilent {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$Attempts = 3
    )

    Remove-Item $OutFile -Force -ErrorAction SilentlyContinue

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
            Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }
    }
    return $false
}

function Get-RemoteSha256Silent {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$Attempts = 3
    )

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            # Use Invoke-WebRequest for 5.1 consistency, keep it quiet
            $r = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }

            $txt = ($r.Content | Out-String)
            if ([string]::IsNullOrWhiteSpace($txt)) { throw "empty" }

            # take first non-empty line
            $line = ($txt -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1).Trim()
            if ([string]::IsNullOrWhiteSpace($line)) { throw "blank" }

            # accept "HASH" or "HASH  filename"
            $hash = ($line -split '\s+')[0].Trim().ToUpper()

            if ($hash -match '^[0-9A-F]{64}$') { return $hash }
            throw "invalid"
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
# Execution (ALL SILENT)
# --------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

try { New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null } catch { return }
try { Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

# Download ZIP
if (-not (Invoke-DownloadSilent -Uri $ZipUrl -OutFile $ZipPath -Attempts 3)) { return }

# Hash verification from GitHub
if (-not $SkipHashCheck) {
    $expected = Get-RemoteSha256Silent -Uri $HashUrl -Attempts 3
    if ([string]::IsNullOrWhiteSpace($expected)) {
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }

    try {
        $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
        if ($actual -ne $expected) {
            Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            return
        }
    } catch {
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }
}

# Extract
try { Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force } catch { return }

# Ensure destination
try { New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null } catch { return }

# Normalize ZIP structure (handle nested folder in ZIP)
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
