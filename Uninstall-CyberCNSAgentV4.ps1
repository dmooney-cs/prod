# ==========================================
#   ConnectSecure Agent V4 - Uninstall Tool
# ==========================================

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$hn = $env:COMPUTERNAME
$log = @()
$export = "C:\Script-Export"
$scriptTemp = "C:\Script-Temp"
$batPath = "$scriptTemp\uninstall.bat"

if (-not (Test-Path $export)) { New-Item -Path $export -ItemType Directory | Out-Null }
if (-not (Test-Path $scriptTemp)) { New-Item -Path $scriptTemp -ItemType Directory | Out-Null }

$logFile = "$export\AgentUninstallLog_$ts`_$hn.csv"
$txtFile = "$export\AgentUninstallTranscript_$ts`_$hn.txt"

Start-Transcript -Path $txtFile -Append

# Step 0: Prepare uninstall.bat
Write-Host "`n🧪 Step 0: Preparing uninstall batch file..." -ForegroundColor Cyan
$batUrl = "https://example.com/uninstall.bat"  # Replace if needed
$batContent = $null

try {
    Write-Host "🌐 Attempting to fetch uninstall.bat from: $batUrl" -ForegroundColor Gray
    $batContent = Invoke-WebRequest -Uri $batUrl -UseBasicParsing -TimeoutSec 10
    $batContent = $batContent.Content
    Write-Host "✅ Successfully fetched uninstall.bat from remote." -ForegroundColor Green
} catch {
    Write-Host "⚠️  Failed to fetch from URL. Using cached batch content." -ForegroundColor Yellow
    $batContent = @'
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
}

Set-Content -Path $batPath -Value $batContent -Encoding ASCII -Force
Write-Host "🗂️  uninstall.bat written to disk at $batPath" -ForegroundColor DarkGray
$log += [PSCustomObject]@{ Step = "Prepare uninstall.bat"; Status = "Complete"; Path = $batPath; Time = $ts }

# Step 1: Default Uninstall
Write-Host "`n🧪 Step 1 of 2: Running Default Uninstall Commands..." -ForegroundColor Cyan
try {
    Stop-Service -Name CyberCNSAgent -Force -ErrorAction SilentlyContinue
    Stop-Service -Name ConnectSecureAgentMonitor -Force -ErrorAction SilentlyContinue
    Write-Host "✅ Services stopped." -ForegroundColor Green
} catch {
    Write-Host "⚠️  Service stop error: $_" -ForegroundColor Yellow
}

$regKeys = @(
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ConnectSecure Agent",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ConnectSecure Agent"
)
foreach ($key in $regKeys) {
    try {
        Remove-Item -Path $key -Recurse -Force -ErrorAction Stop
        Write-Host "🧹 Removed registry: $key" -ForegroundColor Green
    } catch {
        Write-Host "⚠️  Could not remove registry: $key" -ForegroundColor Yellow
    }
}
$log += [PSCustomObject]@{ Step = "Default Uninstall"; Status = "Completed"; Time = (Get-Date) }

# Step 2: Advanced Uninstall (inline execution)
Write-Host "`n🧪 Step 2 of 2: Executing Advanced Uninstall Script..." -ForegroundColor Cyan
try {
    & $batPath
    Write-Host "✅ uninstall.bat executed inline." -ForegroundColor Green
    $log += [PSCustomObject]@{ Step = "Advanced Uninstall"; Status = "Executed Inline"; File = $batPath; Time = (Get-Date) }
} catch {
    Write-Host "❌ Failed to execute uninstall.bat: $_" -ForegroundColor Red
    $log += [PSCustomObject]@{ Step = "Advanced Uninstall"; Status = "Failed"; Time = (Get-Date) }
}

$log | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
Write-Host "`n📁 Log exported to: $logFile" -ForegroundColor Cyan
Stop-Transcript
Pause
