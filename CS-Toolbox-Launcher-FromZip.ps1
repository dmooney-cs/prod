# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Toolbox – Download + Launch from ZIP                 ║
# ║ Version: 1.1 | Updated: 2025-08-07                          ║
# ╚═════════════════════════════════════════════════════════════╝

$zipUrl = "https://github.com/dmooney-cs/prod/raw/main/prod-01-01.zip"
$zipPath = "$env:TEMP\cs-toolbox.zip"
$extractPath = "C:\CS-Toolbox-TEMP"
$launcherScript = "CS-Toolbox-Launcher.ps1"

function Pause-Enter($msg = "Press Enter to continue...") {
    Write-Host ""
    Read-Host $msg | Out-Null
}

# Prompt user
Clear-Host
Write-Host "====================================================="
Write-Host "     🧰 ConnectSecure Toolbox Downloader & Launcher"
Write-Host "====================================================="
Write-Host ""
Pause-Enter "Press Enter to download toolbox from GitHub..."

# Download ZIP
try {
    Write-Host "`n⬇ Downloading ZIP from GitHub..."
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Write-Host "✅ Downloaded to: $zipPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to download ZIP: $_" -ForegroundColor Red
    Pause-Enter
    exit
}

# Extract ZIP
try {
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    Write-Host "✅ Extracted to: $extractPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to extract ZIP: $_" -ForegroundColor Red
    Pause-Enter
    exit
}

# Launch toolbox
$launcherPath = Join-Path $extractPath $launcherScript

if (Test-Path $launcherPath) {
    Pause-Enter "`nPress Enter to launch the toolbox..."
    Write-Host "`n🚀 Launching: $launcherPath`n"
    Set-Location $extractPath
    . $launcherPath  # runs inline in same window
} else {
    Write-Host "❌ Launcher script not found: $launcherPath" -ForegroundColor Red
    Pause-Enter
    exit
}
