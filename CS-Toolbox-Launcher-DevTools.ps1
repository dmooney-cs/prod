# CS-Toolbox-Launcher-FromZip.ps1
# Bootstrapper for ConnectSecure Technician Toolbox (prod-01-01) with SHA-256 verification
# - Downloads toolbox.zip (with retry logic)
# - Verifies SHA-256 against a pinned value ($ExpectedHash)
# - Extracts to C:\CS-Toolbox-TEMP\prod-01-01
# - Launches CS-Toolbox-Launcher.ps1 in the SAME PowerShell window (dot-sourced)
# - Handles nested folders in the ZIP and unblocks files

param(
    [switch]$AutoYes,      # Skip the Y/N prompt
    [switch]$SkipHashCheck # (Not recommended) skip SHA-256 verification
)

# --------------------------
# Config
# --------------------------
$ZipUrl       = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath      = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath  = 'C:\CS-Toolbox-TEMP'
$DestRoot     = Join-Path $ExtractPath 'prod-01-01'
$Launcher     = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

# Pinned, known-good SHA-256 hash for the ZIP (64 hex chars)
$ExpectedHash = 'AF1812EB00FCA40C4FCB564107F21D3E56B044C9C1A34F97FF0ECE37350AC197'

# --------------------------
# Prompt user
# --------------------------
if (-not $AutoYes) {
    $response = Read-Host 'Download and install the ConnectSecure Technician Toolbox (prod-01-01)? (Y/N)'
    if ($response -notin @('Y','y')) {
        Write-Host 'Aborted by user.' -ForegroundColor Yellow
        return
    }
}

# --------------------------
# Prep environment
# --------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

# Ensure base folder exists
if (-not (Test-Path -LiteralPath $ExtractPath)) {
    try {
        New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
    } catch {
        Write-Host ('❌ ERROR: Failed to create {0}: {1}' -f $ExtractPath, $_.Exception.Message) -ForegroundColor Red
        return
    }
}

# Clean existing destination (avoid stale files)
if (Test-Path -LiteralPath $DestRoot) {
    try {
        Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Host ('⚠️ WARN: Could not remove existing folder {0}: {1}' -f $DestRoot, $_.Exception.Message) -ForegroundColor Yellow
    }
}

# --------------------------
# Helpers
# --------------------------
function Invoke-DownloadWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )
    if (Test-Path -LiteralPath $OutFile) {
        Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
    }
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Host ("Downloading: {0} (Attempt {1}/{2})" -f $Uri, $attempt, $MaxAttempts) -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
            if ((Test-Path -LiteralPath $OutFile) -and ((Get-Item -LiteralPath $OutFile).Length -gt 0)) {
                Write-Host "✅ Download successful." -ForegroundColor Green
                return $true
            } else {
                throw "Downloaded file is missing or empty."
            }
        } catch {
            if (Test-Path -LiteralPath $OutFile) {
                Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
            }
            if ($attempt -eq $MaxAttempts) {
                Write-Host ("❌ ERROR: Download failed on attempt {0}/{1}: {2}" -f $attempt, $MaxAttempts, $_.Exception.Message) -ForegroundColor Red
                return $false
            } else {
                Write-Host ("⚠️ Attempt {0}/{1} failed: {2}" -f $attempt, $MaxAttempts, $_.Exception.Message) -ForegroundColor Yellow
                Write-Host ("   Retrying in {0} seconds..." -f $DelaySeconds) -ForegroundColor Yellow
                Start-Sleep -Seconds $DelaySeconds
            }
        }
    }
    return $false
}

function Get-ExpectedSha256 {
    param(
        [string]$PinnedHash
    )
    if (-not [string]::IsNullOrWhiteSpace($PinnedHash)) {
        return $PinnedHash.Trim().ToUpper()
    }
    return $null
}

function Move-Contents([string]$Source, [string]$Target) {
    if (-not (Test-Path -LiteralPath $Source)) { return }
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        try { Move-Item -LiteralPath $_.FullName -Destination $Target -Force } catch {
            Write-Host ('⚠️ WARN: Failed to move {0} -> {1}: {2}' -f $_.FullName, $Target, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# --------------------------
# Download ZIP
# --------------------------
$downloadOk = Invoke-DownloadWithRetry -Uri $ZipUrl -OutFile $ZipPath -MaxAttempts 3 -DelaySeconds 2
if (-not $downloadOk) {
    Write-Host 'Download did not succeed after 3 attempts. Please check your connection or try again later.' -ForegroundColor Red
    return
}

# --------------------------
# SHA-256 Verification (Pinned Only)
# --------------------------
if (-not $SkipHashCheck) {
    $expected = Get-ExpectedSha256 -PinnedHash $ExpectedHash
    if ([string]::IsNullOrWhiteSpace($expected)) {
        Write-Host "⚠️ No expected SHA-256 configured. Aborting to be safe." -ForegroundColor Yellow
        Write-Host "   Tip: set `$ExpectedHash in the script to a known-good hash." -ForegroundColor Yellow
        try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
        return
    }

    try {
        $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
        if ($actual -ne $expected) {
            Write-Host "❌ Hash mismatch! ZIP will be discarded." -ForegroundColor Red
            Write-Host ("   Expected: {0}" -f $expected) -ForegroundColor Red
            Write-Host ("   Actual  : {0}" -f $actual) -ForegroundColor Red
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            return
        } else {
            Write-Host "✅ File hash verified (SHA-256)." -ForegroundColor Green
        }
    } catch {
        Write-Host ("❌ ERROR computing SHA-256: {0}" -f $_.Exception.Message) -ForegroundColor Red
        try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
        return
    }
} else {
    Write-Host "⚠️ Hash verification skipped by user switch." -ForegroundColor Yellow
}

# --------------------------
# Extract ZIP
# --------------------------
Write-Host 'Extracting toolbox...' -ForegroundColor Cyan
try {
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
} catch {
    Write-Host ('❌ ERROR: Extract failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
    return
}

# Ensure destination exists for normalization
if (-not (Test-Path -LiteralPath $DestRoot)) {
    New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null
}

# Normalize folder structure
$alreadyGood = Test-Path -LiteralPath (Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1')
if (-not $alreadyGood) {
    $topDirs = Get-ChildItem -LiteralPath $ExtractPath -Directory -Force | Where-Object { $_.FullName -ne $DestRoot }
    $topFiles = Get-ChildItem -LiteralPath $ExtractPath -File -Force | Where-Object { $_.FullName -ne $ZipPath }

    if ($topDirs.Count -eq 1 -and $topFiles.Count -eq 0) {
        Move-Contents -Source $topDirs[0].FullName -Target $DestRoot
        try { Remove-Item -LiteralPath $topDirs[0].FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    } else {
        foreach ($d in $topDirs) { Move-Contents -Source $d.FullName -Target $DestRoot }
        foreach ($f in $topFiles) { try { Move-Item -LiteralPath $f.FullName -Destination $DestRoot -Force } catch { } }
        foreach ($d in $topDirs) { try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { } }
    }
}

# Unblock all extracted files
try {
    Get-ChildItem -LiteralPath $DestRoot -Recurse -Force -File | ForEach-Object {
        try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { }
    }
} catch { }

# Verify launcher exists
if (-not (Test-Path -LiteralPath $Launcher)) {
    Write-Host ('❌ ERROR: Launcher not found: {0}' -f $Launcher) -ForegroundColor Red
    Write-Host 'Please verify the ZIP contents or try again.' -ForegroundColor Yellow
    return
}

# --------------------------
# Ready to launch in SAME window
# --------------------------
Write-Host '✅ Download & extraction complete.' -ForegroundColor Green
if (-not $AutoYes) { $null = Read-Host 'Press ENTER to launch the ConnectSecure Technician Toolbox' }

try {
    . $Launcher   # dot-source; keep in same PowerShell session
} catch {
    Write-Host ('❌ ERROR launching Toolbox: {0}' -f $_.Exception.Message) -ForegroundColor Red
    if (-not $AutoYes) { $null = Read-Host 'Press ENTER to exit' }
}
