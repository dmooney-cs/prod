# ╔════════════════════════════════════════════════════════════════╗
# ║  🚀 ConnectSecure Technician Toolbox Bootstrap Script          ║
# ║  Version: 1.0 | Downloads and launches the CS Toolbox          ║
# ╚════════════════════════════════════════════════════════════════╝

$zipUrl = "https://github.com/dmooney-cs/prod/raw/main/cs-toolbox-v1-0.zip"
$zipPath = "$env:TEMP\cs-toolbox.zip"
$extractPath = "C:\CS-Toolbox-TEMP"

# Prompt to begin
Read-Host "Press ENTER to download the ConnectSecure Technician Toolbox..."

# Download ZIP
Write-Host "`n📥 Downloading toolbox..." -ForegroundColor Cyan
Start-BitsTransfer -Source $zipUrl -Destination $zipPath

# Confirm file exists with timeout
$timeout = 0
while (!(Test-Path $zipPath) -and $timeout -lt 10) {
    Start-Sleep -Seconds 1
    $timeout++
}

if (!(Test-Path $zipPath)) {
    Write-Host "❌ Download failed within timeout. Exiting." -ForegroundColor Red
    exit 1
}

# Extract contents
Write-Host "📦 Extracting toolbox to $extractPath..." -ForegroundColor Cyan
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

# Locate launcher script
$launcher = Get-ChildItem -Path $extractPath -Filter "CS-Toolbox-Launcher.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $launcher) {
    Write-Host "❌ Launcher script not found. Please check the extracted folder." -ForegroundColor Red
    exit 1
}

# Prompt before launch
Read-Host "`n✅ Toolbox is ready. Press ENTER to launch the ConnectSecure Technician Toolbox..."

# Launch toolbox
Start-Process powershell -ArgumentList "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $launcher.FullName