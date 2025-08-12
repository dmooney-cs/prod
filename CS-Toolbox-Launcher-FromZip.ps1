# Get-CS-Toolbox.ps1 — bootstrap for ConnectSecure Technician Toolbox (prod-01-01.zip)

$zipUrl      = "https://github.com/dmooney-cs/prod/raw/main/prod-01-01.zip"  # RAW file url (not /blob/)
$zipPath     = Join-Path $env:TEMP "prod-01-01.zip"
$extractPath = "C:\CS-Toolbox-TEMP"
$destRoot    = Join-Path $extractPath "prod-01-01"

# Ask before we do anything
$response = Read-Host "Download and install the ConnectSecure Technician Toolbox (prod-01-01)? (Y/N)"
if ($response.ToUpper() -ne "Y") {
    Write-Host "Aborted by user." -ForegroundColor Yellow
    exit
}

# Ensure TLS 1.2 for web requests (older hosts)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Download (BITS first; fallback to Invoke-WebRequest)
Write-Host "Downloading toolbox..." -ForegroundColor Cyan
$downloaded = $false
try {
    Start-BitsTransfer -Source $zipUrl -Destination $zipPath -ErrorAction Stop
    $downloaded = $true
} catch {
    try {
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        $downloaded = $true
    } catch {
        Write-Host "❌ Download failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

if (-not (Test-Path $zipPath)) {
    Write-Host "❌ ZIP not found after download. Exiting." -ForegroundColor Red
    exit 1
}

# Prep destination
if (Test-Path $destRoot) {
    Write-Host "Clearing previous install at $destRoot ..." -ForegroundColor DarkCyan
    try { Remove-Item -Path $destRoot -Recurse -Force -ErrorAction Stop } catch {}
}
if (-not (Test-Path $extractPath)) { New-Item -Path $extractPath -ItemType Directory -Force | Out-Null }
New-Item -Path $destRoot -ItemType Directory -Force | Out-Null

# Extract to C:\CS-Toolbox-TEMP
Write-Host "Extracting toolbox..." -ForegroundColor Cyan
try {
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
} catch {
    Write-Host "❌ Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# If content landed under a single top-level folder, flatten into C:\CS-Toolbox-TEMP\prod-01-01
$topItems = Get-ChildItem -Path $extractPath -Force | Where-Object { $_.Name -ne "prod-01-01" -and $_.Name -ne (Split-Path $zipPath -Leaf) }
# Prefer a folder named 'prod-01-01' if it already exists; otherwise, detect single-folder nesting
$nestedCandidate = $null
if (-not (Test-Path $destRoot)) {
    # If prod-01-01 wasn't created yet but exists as extracted folder
    $nestedCandidate = Get-ChildItem -Path $extractPath -Directory -Force | Where-Object { $_.Name -match '^prod-01-01$' } | Select-Object -First 1
    if ($nestedCandidate) { New-Item -Path $destRoot -ItemType Directory -Force | Out-Null }
}
if (-not $nestedCandidate) {
    $onlyDir = Get-ChildItem -Path $extractPath -Directory -Force | Where-Object { $_.FullName -ne $destRoot } | Select-Object -First 1
    if ($onlyDir -and (Get-ChildItem -Path $extractPath -Directory -Force).Count -eq 1) { $nestedCandidate = $onlyDir }
}

if ($nestedCandidate) {
    # Move everything from nested folder into prod-01-01
    Get-ChildItem -Path $nestedCandidate.FullName -Force | Move-Item -Destination $destRoot -Force
    # Remove the now-empty nested folder
    try { Remove-Item -Path $nestedCandidate.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
} else {
    # If files dumped directly in extractPath, move them into prod-01-01
    Get-ChildItem -Path $extractPath -Force | Where-Object { $_.FullName -ne $destRoot -and $_.Name -ne (Split-Path $zipPath -Leaf) } | `
        Move-Item -Destination $destRoot -Force
}

# Unblock scripts so they run without SmartScreen hassles
Get-ChildItem -Path $destRoot -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
    try { Unblock-File -Path $_.FullName } catch {}
}

# Find launcher and run
$launcher = Get-ChildItem -Path $destRoot -Filter "CS-Toolbox-Launcher.ps1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $launcher) {
    Write-Host "❌ Launcher not found after extraction." -ForegroundColor Red
    Write-Host "Looked in: $destRoot"
    exit 1
}

Read-Host "✅ Download & extraction complete. Press ENTER to launch the ConnectSecure Technician Toolbox"
Start-Process powershell -ArgumentList "-NoLogo","-NoProfile","-ExecutionPolicy","Bypass","-File","`"$($launcher.FullName)`""
