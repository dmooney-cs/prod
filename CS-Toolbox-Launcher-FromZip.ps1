# ╔════════════════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Toolbox – Download + Launch from ZIP                             ║
# ║ Version: 1.4 | ExecutionPolicy Bypass Enabled                          ║
# ╚════════════════════════════════════════════════════════════════════════╝

$zipUrl = "https://github.com/dmooney-cs/prod/raw/main/prod-01-01.zip"
$zipPath = "$env:TEMP\cs-toolbox.zip"
$extractPath = "C:\CS-Toolbox-TEMP"
$launcherPath = "C:\CS-Toolbox-TEMP\prod-01-01\CS-Toolbox-Launcher.ps1"

function Pause-Enter($msg = "Press Enter to continue...") {
    Write-Host ""
    Read-Host $msg | Out-Null
}

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
    if (Test-Path "$extractPath\prod-01-01") {
        Remove-Item "$extractPath\prod-01-01" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    Write-Host "✅ Extracted to: $extractPath\prod-01-01" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to extract ZIP: $_" -ForegroundColor Red
    Pause-Enter
    exit
}

# Step 3: Launch new process with ExecutionPolicy Bypass
if (Test-Path $launcherPath) {
    Pause-Enter "`nPress Enter to launch the CS Toolbox..."
    Write-Host "`n🚀 Launching with ExecutionPolicy Bypass...`n"

    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoLogo -NoExit -File `"$launcherPath`""
} else {
    Write-Host "❌ Launcher script not found at:`n$launcherPath" -ForegroundColor Red
    Pause-Enter
    exit
}
