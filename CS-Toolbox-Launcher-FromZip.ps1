# ================================================================
# ConnectSecure Toolbox Bootstrap (Run-ConnectSecure-Toolbox.ps1)
# Downloads ZIP, extracts to C:\CS-Toolbox-TEMP, launches launcher
# Safe for Windows PowerShell 5.1
# ================================================================

# 0) Make this session permissive so no prompts stop us
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}

# 1) Networking hardening (TLS 1.2+)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} catch {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

# 2) Inputs / paths
$zipUrl      = "https://github.com/dmooney-cs/prod/raw/main/cs-toolbox-v1-0.zip"
$zipPath     = Join-Path $env:TEMP "cs-toolbox.zip"
$extractPath = "C:\CS-Toolbox-TEMP"

# 3) Confirm
Write-Host "This will download the ConnectSecure Technician Toolbox to:" -ForegroundColor Cyan
Write-Host "  $extractPath" -ForegroundColor Gray
$null = Read-Host "Press ENTER to continue (or Ctrl+C to cancel)"

# 4) Prep folder
try {
    if (-not (Test-Path $extractPath)) { New-Item -Path $extractPath -ItemType Directory -Force | Out-Null }
} catch {
    Write-Host "❌ Unable to create $extractPath : $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 5) Download (BITS first, fallback to Invoke-WebRequest)
Write-Host "Downloading toolbox..." -ForegroundColor Cyan
$downloadOk = $false
try {
    Import-Module BitsTransfer -ErrorAction SilentlyContinue | Out-Null
    Start-BitsTransfer -Source $zipUrl -Destination $zipPath -ErrorAction Stop
    $downloadOk = $true
} catch {
    Write-Host "BITS failed, trying Invoke-WebRequest..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        $downloadOk = $true
    } catch {
        Write-Host "❌ Download failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# 6) Verify download appeared (with brief wait)
$deadline = (Get-Date).AddSeconds(10)
while (-not (Test-Path $zipPath) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
if (-not (Test-Path $zipPath)) {
    Write-Host "❌ Download did not produce a file at $zipPath" -ForegroundColor Red
    exit 1
}

# 7) Unblock the ZIP (avoid Mark-of-the-Web)
try { Unblock-File -Path $zipPath -ErrorAction SilentlyContinue } catch {}

# 8) Extract
Write-Host "Extracting toolbox..." -ForegroundColor Cyan
try {
    if (Test-Path $extractPath) {
        # Keep folder, but clear previous extracted content if desired:
        # Remove-Item -Path (Join-Path $extractPath '*') -Recurse -Force -ErrorAction SilentlyContinue
        # (Leaving existing files can help with incremental updates; adjust as you prefer.)
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
} catch {
    Write-Host "❌ Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# 9) Unblock all extracted files to suppress SmartScreen/zone prompts
try {
    Get-ChildItem -Path $extractPath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
        try { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue } catch {}
    }
} catch {}

# 10) Locate a launcher (support both names)
$launcher = Get-ChildItem -Path $extractPath -Recurse -Include "CS-Toolbox-Launcher-FromZip.ps1","CS-Toolbox-Launcher.ps1" -File -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $launcher) {
    Write-Host "❌ Launcher not found after extraction." -ForegroundColor Red
    Write-Host "   Looked for: CS-Toolbox-Launcher-FromZip.ps1 or CS-Toolbox-Launcher.ps1" -ForegroundColor Yellow
    exit 1
}

# 11) Summary + launch
Write-Host ""
Write-Host "✅ Download & Extract complete." -ForegroundColor Green
Write-Host ("   ZIP:       {0}" -f $zipPath) -ForegroundColor Gray
Write-Host ("   Extracted: {0}" -f $extractPath) -ForegroundColor Gray
Write-Host ("   Launcher:  {0}" -f $launcher.FullName) -ForegroundColor Gray
$null = Read-Host "Press ENTER to launch the ConnectSecure Technician Toolbox"

# Always launch with ExecutionPolicy Bypass in a new window
try {
    Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoLogo","-NoProfile",
        "-ExecutionPolicy","Bypass",
        "-File", "`"$($launcher.FullName)`""
    ) -Verb Open | Out-Null
} catch {
    Write-Host "❌ Failed to start launcher: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
