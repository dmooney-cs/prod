# CS-Toolbox-Launcher-DevTools-ZeroTouch.ps1
# Silent by default.
# If -ShowHashes or -ConfirmHashes is used:
#   - shows stage-by-stage status
#   - shows expected hash (GitHub) + actual zip hash
#   - shows failure reason if it exits early

param(
    [switch]$SkipHashCheck = $false,
    [switch]$ShowHashes = $false,
    [switch]$ConfirmHashes = $false
)

#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'  # kill progress UI

# --------------------------
# Config
# --------------------------
$ZipUrl      = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath     = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath = 'C:\CS-Toolbox-TEMP'
$DestRoot    = Join-Path $ExtractPath 'prod-01-01'
$Launcher    = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

# IMPORTANT: set this to your real GitHub RAW URL
$HashUrl     = 'https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/devtools.sha256'

# --------------------------
# Output control
# --------------------------
$script:AllowOutput = [bool]($ShowHashes -or $ConfirmHashes)

function Say([string]$msg) {
    if ($script:AllowOutput) { Write-Host $msg }
}

function Fail([string]$stage, [string]$reason) {
    if ($script:AllowOutput) {
        Write-Host ""
        Write-Host "FAILED at: $stage"
        Write-Host "Reason : $reason"
        Write-Host ""
    }
    return $false
}

# --------------------------
# Helpers (quiet)
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
            if ([string]::IsNullOrWhiteSpace($txt)) { throw "empty response" }

            $line = ($txt -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1).Trim()
            if ([string]::IsNullOrWhiteSpace($line)) { throw "blank file" }

            $hash = ($line -split '\s+')[0].Trim().ToUpper()
            if ($hash -match '^[0-9A-F]{64}$') { return $hash }

            throw "not a valid SHA-256 (expected 64 hex chars)"
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
# Main (silent unless switches)
# --------------------------
$stage = "init"

try {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    $stage = "prepare folders"
    try { New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null } catch { return (Fail $stage $_.Exception.Message) }
    try { Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

    $stage = "download zip"
    Say "Downloading ZIP..."
    if (-not (Invoke-DownloadQuiet -Uri $ZipUrl -OutFile $ZipPath -Attempts 3)) {
        return (Fail $stage "Unable to download ZIP from $ZipUrl")
    }

    $expected = $null
    $actual   = $null

    if (-not $SkipHashCheck) {
        $stage = "download expected hash"
        Say "Fetching expected SHA-256 from GitHub..."
        $expected = Get-RemoteSha256Quiet -Uri $HashUrl -Attempts 3
        if ([string]::IsNullOrWhiteSpace($expected)) {
            Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            return (Fail $stage "Unable to download/parse SHA-256 from $HashUrl")
        }

        $stage = "compute zip hash"
        Say "Computing ZIP SHA-256..."
        try {
            $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
        } catch {
            Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            return (Fail $stage $_.Exception.Message)
        }

        if ($ShowHashes -or $ConfirmHashes) {
            Write-Host ""
            Write-Host "Expected (GitHub): $expected"
            Write-Host "Actual   (ZIP)   : $actual"
            Write-Host ""
        }

        $stage = "compare hashes"
        if ($actual -ne $expected) {
            Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            return (Fail $stage "Hash mismatch")
        }

        if ($ConfirmHashes) {
            $resp = Read-Host "Hashes match. Proceed with extract + launch? (Y/N)"
            if ($resp -notin @('Y','y')) { return $true }
        }
    } else {
        if ($ConfirmHashes -or $ShowHashes) {
            Write-Host ""
            Write-Host "Hash check skipped (-SkipHashCheck)."
            Write-Host ""
        }
        if ($ConfirmHashes) {
            $resp = Read-Host "Proceed without hash verification? (Y/N)"
            if ($resp -notin @('Y','y')) { return $true }
        }
    }

    $stage = "extract zip"
    Say "Extracting..."
    try { Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force } catch { return (Fail $stage $_.Exception.Message) }

    $stage = "ensure dest"
    try { New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null } catch { return (Fail $stage $_.Exception.Message) }

    $stage = "normalize structure"
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

    $stage = "unblock"
    try {
        Get-ChildItem -LiteralPath $DestRoot -Recurse -Force -File |
            ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { } }
    } catch { }

    $stage = "launch"
    if (-not (Test-Path -LiteralPath $Launcher)) {
        return (Fail $stage "Launcher not found at $Launcher")
    }

    Say "Launching..."
    . $Launcher
    return $true

} catch {
    return (Fail $stage $_.Exception.Message)
}
