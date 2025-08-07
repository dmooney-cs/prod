# ╔════════════════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Toolbox – Download + Launch from ZIP                             ║
# ║ Version: 1.2 | Updated: 2025-08-07 | Extracts & launches v01-01 build  ║
# ╚════════════════════════════════════════════════════════════════════════╝

$zipUrl = "https://github.com/dmooney-cs/prod/raw/main/prod-01-01.zip"
$zipPath = "$env:TEMP\cs-toolbox.zip"
$extractRoot = "C:\CS-Toolbox-TEMP"
$extractSubfolder = "prod-01-01"
$fullExtractPath = Join-Path $extractRoot $extractSubfolder
$launcherScript = "CS-Toolbox-Launcher.ps1"

function Pause-Enter($msg = "Press Enter to continue...") {
    Write-Host ""
    Read-Host $msg | Out-Null
}

# UI Header
Clear-Host
Write-Host "====================================================="
Write-Host "     🧰 ConnectSecure Toolbox Downloader & Launcher"
Write-Host "====================================================="
Write-Host ""
Pause-Enter "Press Enter to download the toolbox from GitHub..."

# Step 1: Download ZIP
try {
    Write-Host "`n⬇ Downloading toolbox ZIP from GitHub..."
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Write-Host "✅ Downloaded to: $zipPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to download ZIP: $_" -ForegroundColor Red
    Pause-Enter
    exit
}

# Step 2: Extract ZIP
try {
    if (Test-Path $fullExtractPath) {
        Remove-Item $fullExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractRoot -Force
    Write-Host "✅ Extracted to: $fullExtractPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to extract ZIP: $_" -ForegroundColor Red
    Pause-Enter
    exit
}

# Step 3: Launch toolbox inline
$launcherPath = Join-Path $fullExtractPath $launcherScript

if (Test-Path $launcherPath) {
    Pause-Enter "`nPress Enter to launch the CS Toolbox..."
    Write-Host "`n🚀 Launching: $launcherPath`n"
    Set-Location $fullExtractPath
    . $launcherPath  # run inline in same PowerShell window
} else {
    Write-Host "❌ Launcher script not found at: $launcherPath" -ForegroundColor Red
    Pause-Enter
    exit
}
