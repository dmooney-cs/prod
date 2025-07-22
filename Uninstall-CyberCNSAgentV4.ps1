# ╔══════════════════════════════════════════════════════════════════╗
# ║ 🔧 Uninstall CyberCNS Agent V4 - Full Script (With Fallback)     ║
# ╚══════════════════════════════════════════════════════════════════╝

. { iwr -useb "https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1" } | iex
Show-Header "CyberCNS Agent V4 Uninstaller"

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$hn = $env:COMPUTERNAME
$log = @()
$folder = Ensure-ExportFolder
$csvPath = "$folder\UninstallAgent_$ts`_$hn.csv"
$txtPath = "$folder\UninstallAgent_$ts`_$hn.txt"
Start-Transcript -Path $txtPath -Append

# Step 1: Default Uninstall
Write-Host "`n🧹 Step 1 of 2: Running Default Uninstall Command..." -ForegroundColor Cyan
try {
    Set-Location "C:\Program Files (x86)\CyberCNSAgent"
    Write-Host "Executing: cybercnsagent.exe -r" -ForegroundColor Yellow
    .\cybercnsagent.exe -r
    $log += [PSCustomObject]@{ Step = "Default Uninstall"; Status = "Executed"; Time = $ts }
} catch {
    Write-Host "❌ Default uninstall failed: $_" -ForegroundColor Red
    $log += [PSCustomObject]@{ Step = "Default Uninstall"; Status = "Failed: $_"; Time = $ts }
}

# Step 2: Advanced Batch Uninstall
Write-Host "`n🧪 Step 2 of 2: Running Advanced Uninstall Script..." -ForegroundColor Cyan
$batUrl = "https://example.com/uninstall.bat"  # Replace with real URL
$batPath = "C:\Program Files (x86)\CyberCNSAgent\uninstall.bat"
$batContent = ""

try {
    Write-Host "🌐 Attempting to fetch uninstall.bat from: $batUrl" -ForegroundColor Yellow
    $batContent = Invoke-WebRequest -Uri $batUrl -UseBasicParsing -TimeoutSec 10
    $batContent.Content | Out-File -Encoding ASCII -FilePath $batPath -Force
    Write-Host "✅ Downloaded uninstall.bat from online source." -ForegroundColor Green
    $log += [PSCustomObject]@{ Step = "Fetch uninstall.bat"; Status = "Downloaded from URL"; Time = $ts }
} catch {
    Write-Host "⚠️  Failed to fetch from URL. Using cached batch content." -ForegroundColor DarkYellow
    $cached = @'
@echo off
ping 127.0.0.1 -n 6 > nul
cd "C:\PROGRA~2"
sc stop ConnectSecureAgentMonitor
timeout /T 5 > nul
sc delete ConnectSecureAgentMonitor
timeout /T 5 > nul
sc stop CyberCNSAgent
timeout /T 5 > nul
sc delete CyberCNSAgent
ping 127.0.0.1 -n 6 > nul
taskkill /IM osqueryi.exe /F
taskkill /IM nmap.exe /F
taskkill /IM cyberutilities.exe /F
CyberCNSAgent\cybercnsagent.exe --internalAssetArgument uninstallservice
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ConnectSecure Agent" /f
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ConnectSecure Agent" /f
rmdir CyberCNSAgent /s /q
'@
    $cached | Out-File -FilePath $batPath -Encoding ASCII -Force
    Write-Host "🗂️  Cached uninstall.bat written to disk." -ForegroundColor Green
    $log += [PSCustomObject]@{ Step = "Fetch uninstall.bat"; Status = "Fallback to cached version"; Time = $ts }
}

# Execute uninstall.bat
try {
    Write-Host "`n🚀 Executing uninstall.bat..." -ForegroundColor Cyan
    cmd.exe /c "`"$batPath`""
    Write-Host "✅ Advanced uninstall batch script completed from: $batPath" -ForegroundColor Green
    $log += [PSCustomObject]@{ Step = "Run uninstall.bat"; Status = "Executed from $batPath"; Time = $ts }
} catch {
    Write-Host "❌ Error running uninstall.bat: $_" -ForegroundColor Red
    $log += [PSCustomObject]@{ Step = "Run uninstall.bat"; Status = "Failed: $_"; Time = $ts }
}

# Export logs
Export-Data -Data $log -Path $csvPath
Write-Host "`n📁 CSV log exported to: " -NoNewline; Write-ExportPath $csvPath
Write-Host "📄 Transcript saved to: " -NoNewline; Write-ExportPath $txtPath

# Zip and Email Results
Run-ZipAndEmailResults

# Cleanup
Run-CleanupExportFolder

Pause-Script "Uninstall routine complete. Press any key to close."
Stop-Transcript
