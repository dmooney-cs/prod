# CS-Toolbox-DownloadOnly.ps1
# Download-only bootstrapper for ConnectSecure Technician Toolbox (prod-01-01)
# - Downloads prod-01-01.zip (with retry logic)
# - Verifies SHA-256 against a pinned value ($ExpectedHash)  [REQUIRED]
# - Extracts to C:\CS-Toolbox-TEMP\prod-01-01
# - DOES NOT launch the CS-Toolbox-Launcher.ps1
# - Automatically sets the working directory to the toolbox root
# - Waits for required files to exist before returning (prevents cookbook races)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# --------------------------
# Config
# --------------------------
$ZipUrl       = 'https://github.com/dmooney-cs/prod/raw/refs/heads/main/prod-01-01.zip'
$ZipPath      = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath  = 'C:\CS-Toolbox-TEMP'
$DestRoot     = Join-Path $ExtractPath 'prod-01-01'
$Launcher     = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

# Pinned known-good SHA-256 (REQUIRED)
$ExpectedHash = 'AF1812EB00FCA40C4FCB564107F21D3E56B044C9C1A34F97FF0ECE37350AC197'

# --------------------------
# Wait Helpers
# --------------------------
function Wait-Path {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSec = 180,
        [int]$PollMs = 250
    )
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath $Path)) {
        if ($sw.Elapsed.TotalSeconds -ge $TimeoutSec) {
            throw "Timed out waiting for: $Path"
        }
        Start-Sleep -Milliseconds $PollMs
    }
}

function Wait-Files {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string[]]$Files,
        [int]$TimeoutSec = 180
    )
    foreach ($f in $Files) {
        Wait-Path -Path (Join-Path $Root $f) -TimeoutSec $TimeoutSec
    }
}

# --------------------------
# Prep environment
# --------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Ensure base folder exists
if (-not (Test-Path -LiteralPath $ExtractPath)) {
    try {
        New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
    } catch {
        Write-Host "❌ ERROR: Failed to create $ExtractPath : $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Clean existing destination
if (Test-Path -LiteralPath $DestRoot) {
    try { Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction Stop }
    catch {
        Write-Host "⚠️ WARN: Could not remove existing folder $DestRoot : $($_.Exception.Message)" -ForegroundColor Yellow
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

    if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Host "Downloading: $Uri (Attempt $attempt/$MaxAttempts)" -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
            if ((Test-Path -LiteralPath $OutFile) -and ((Get-Item -LiteralPath $OutFile).Length -gt 0)) {
                Write-Host "✅ Download successful." -ForegroundColor Green
                return $true
            }
            throw "Downloaded file is empty."
        } catch {
            if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }
            if ($attempt -eq $MaxAttempts) {
                Write-Host "❌ ERROR: Download failed: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
            Write-Host "⚠️ Attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "   Retrying in $DelaySeconds seconds..."
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $false
}

function Move-Contents([string]$Source, [string]$Target) {
    if (-not (Test-Path -LiteralPath $Source)) { return }
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        try { Move-Item -LiteralPath $_.FullName -Destination $Target -Force }
        catch {
            Write-Host "⚠️ WARN: Failed to move $($_.FullName) → $Target : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# --------------------------
# Download ZIP
# --------------------------
$downloadOk = Invoke-DownloadWithRetry -Uri $ZipUrl -OutFile $ZipPath
if (-not $downloadOk) { exit 1 }

# --------------------------
# SHA-256 Verification (REQUIRED)
# --------------------------
try {
    $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
    if ($actual -ne $ExpectedHash) {
        Write-Host "❌ Hash mismatch! ZIP discarded." -ForegroundColor Red
        Write-Host "   Expected: $ExpectedHash" -ForegroundColor Red
        Write-Host "   Actual  : $actual" -ForegroundColor Red
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        exit 1
    }
    Write-Host "✅ File hash verified (SHA-256)." -ForegroundColor Green
} catch {
    Write-Host "❌ ERROR computing SHA-256: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --------------------------
# Extract ZIP
# --------------------------
Write-Host "Extracting toolbox..." -ForegroundColor Cyan
try {
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force
} catch {
    Write-Host "❌ ERROR: Extract failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Normalize structure
if (-not (Test-Path -LiteralPath $DestRoot)) { New-Item -LiteralPath $DestRoot -ItemType Directory -Force | Out-Null }

$alreadyGood = Test-Path -LiteralPath (Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1')
if (-not $alreadyGood) {
    $topDirs  = Get-ChildItem -LiteralPath $ExtractPath -Directory -Force | Where-Object { $_.FullName -ne $DestRoot }
    $topFiles = Get-ChildItem -LiteralPath $ExtractPath -File -Force | Where-Object { $_.FullName -ne $ZipPath }

    if ($topDirs.Count -eq 1 -and $topFiles.Count -eq 0) {
        Move-Contents -Source $topDirs[0].FullName -Target $DestRoot
        Remove-Item -LiteralPath $topDirs[0].FullName -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        foreach ($d in $topDirs)  { Move-Contents -Source $d.FullName -Target $DestRoot }
        foreach ($f in $topFiles) { Move-Item -LiteralPath $f.FullName -Destination $DestRoot -Force }
        foreach ($d in $topDirs)  { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# --------------------------
# Wait for required files (prevents downstream races)
# --------------------------
try {
    Wait-Path  -Path $DestRoot -TimeoutSec 300
    Wait-Files -Root $DestRoot -Files @(
        'CS-Toolbox-Launcher.ps1'
        # Add others here if cookbooks rely on them immediately, e.g.:
        # 'Registry-Search.ps1'
    ) -TimeoutSec 300
} catch {
    Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Unblock all extracted files
Get-ChildItem -LiteralPath $DestRoot -Recurse -File -Force | ForEach-Object {
    try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch {}
}

# --------------------------
# Final Status + Auto-CD
# --------------------------
Write-Host "✅ Download & extraction complete (download-only mode)." -ForegroundColor Green
Write-Host "Toolbox root : $DestRoot"  -ForegroundColor Cyan
Write-Host "Launcher path: $Launcher" -ForegroundColor Cyan

try {
    Set-Location -LiteralPath $DestRoot
    Write-Host "📁 Working directory changed to: $DestRoot" -ForegroundColor Green
} catch {
    Write-Host "⚠️ Could not change directory: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Emit object for automation usage
[pscustomobject]@{
    DestRoot = $DestRoot
    Launcher = $Launcher
    ZipPath  = $ZipPath
    HashOk   = $true
}

