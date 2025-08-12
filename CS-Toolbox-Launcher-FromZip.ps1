$zipUrl = "https://github.com/dmooney-cs/prod/raw/main/prod-01-01.zip"
$zipPath = "$env:TEMP\cs-toolbox.zip"
$extractPath = "C:\CS-Toolbox-TEMP"

# Prompt for download
$response = Read-Host "Do you want to download the ConnectSecure Technician Toolbox? (Y/N)"
if ($response.ToUpper() -ne "Y") {
    Write-Host "Aborted by user." -ForegroundColor Yellow
    exit
}

# Start the download
Write-Host "Downloading toolbox..." -ForegroundColor Cyan
Start-BitsTransfer -Source $zipUrl -Destination $zipPath

# Wait up to 10 seconds for the file to appear
$timeout = 0
while (!(Test-Path $zipPath) -and $timeout -lt 10) {
    Start-Sleep -Seconds 1
    $timeout++
}

if (!(Test-Path $zipPath)) {
    Write-Host "❌ Download did not complete within 10 seconds. Exiting." -ForegroundColor Red
    exit 1
}

# Ensure target exists and is empty
if (Test-Path $extractPath) {
    Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -Path $extractPath -ItemType Directory -Force | Out-Null

# Extract contents (flatten if the ZIP has a top-level folder)
Write-Host "Extracting toolbox..." -ForegroundColor Cyan
Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

# If everything got dumped in a subfolder, move it up
$topItems = Get-ChildItem -Path $extractPath
if ($topItems.Count -eq 1 -and $topItems[0].PSIsContainer) {
    Get-ChildItem -Path $topItems[0].FullName -Force | Move-Item -Destination $extractPath -Force
    Remove-Item -Path $topItems[0].FullName -Recurse -Force
}

# Locate launcher recursively
$launcher = Get-ChildItem -Path $extractPath -Filter "CS-Toolbox-Launcher.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $launcher) {
    Write-Host "❌ Launcher not found after extraction. Exiting." -ForegroundColor Red
    exit 1
}

# Prompt to start
Read-Host "✅ Download complete. Press ENTER to launch the ConnectSecure Technician Toolbox"
Start-Process powershell -ArgumentList "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $launcher.FullName
