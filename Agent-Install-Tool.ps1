# ==========================================
#   ConnectSecure Agent Install Utility
# ==========================================

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder
Show-Header "ConnectSecure Agent Install Utility"

function Run-AgentInstall {
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $log = @()
    $baseName = "AgentInstallLog_$ts`_$hn"
    $txtPath = "C:\\Script-Temp\\$baseName.txt"
    if (-not (Test-Path "C:\\Script-Temp")) {
        New-Item -Path "C:\\Script-Temp" -ItemType Directory | Out-Null
    }
    Start-Transcript -Path $txtPath -Append

    # Check for existing services
    $svc1 = Get-Service -Name CyberCNSAgent -ErrorAction SilentlyContinue
    $svc2 = Get-Service -Name ConnectSecureAgentMonitor -ErrorAction SilentlyContinue

    if ($svc1 -or $svc2) {
        Write-Host "`n⚠️ ConnectSecure Agent is already installed." -ForegroundColor Yellow
        $status1 = if ($svc1) { "$($svc1.Status)" } else { "Not Found" }
        $status2 = if ($svc2) { "$($svc2.Status)" } else { "Not Found" }
        Write-Host "Service Status:" -ForegroundColor Gray
        Write-Host "  - CyberCNSAgent: $status1"
        Write-Host "  - ConnectSecureAgentMonitor: $status2"

        $choice = Read-Host "Do you want to uninstall first using the advanced tool? (Y/N)"
        if ($choice -match "^[Yy]$") {
            Write-Host "`n🧪 Step 1/2: Preparing uninstall script..." -ForegroundColor Cyan
            $uninstallPath = "C:\\Script-Temp\\uninstall.bat"
            $sourcePath = "C:\\Program Files (x86)\\CyberCNSAgent\\uninstall.bat"

            if (Test-Path $sourcePath) {
                Copy-Item $sourcePath $uninstallPath -Force
                Write-Host "🗂️ uninstall.bat copied from local agent folder." -ForegroundColor Yellow
            } else {
                try {
                    Invoke-WebRequest -Uri "https://example.com/uninstall.bat" -OutFile $uninstallPath -UseBasicParsing -ErrorAction Stop
                    Write-Host "🗂️ uninstall.bat downloaded from fallback URL." -ForegroundColor Yellow
                } catch {
                    Write-Host "⚠️  Failed to fetch from URL. Using cached batch content." -ForegroundColor Yellow
                    $batContent = @"
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
"@
                    $batContent | Out-File -FilePath $uninstallPath -Encoding ASCII -Force
                    Write-Host "🗂️ uninstall.bat written from local cache." -ForegroundColor Yellow
                }
            }

            Write-Host "🧪 Step 2/2: Running uninstall..." -ForegroundColor Cyan
            cmd.exe /c $uninstallPath
        } else {
            Pause-Script "Install aborted. Press any key to return."
            return
        }
    }

    # Prompt for credentials
    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $key     = Read-Host "Enter Secret Key"

    # Download agent
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $downloadUrl = (Invoke-WebRequest -Uri "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows" -UseBasicParsing).Content
    $agentPath = "C:\\Script-Temp\\cybercnsagent.exe"

    Write-Host "`n⬇️ Downloading agent from:" -ForegroundColor Cyan
    Write-Host $downloadUrl -ForegroundColor Gray
    Invoke-WebRequest -Uri $downloadUrl -OutFile $agentPath -UseBasicParsing

    $size = (Get-Item $agentPath).Length / 1MB
    Write-Host "✅ Agent downloaded to $agentPath ($([math]::Round($size,2)) MB)" -ForegroundColor Green
    $log += [PSCustomObject]@{
        Step     = "Download Agent"
        Status   = "Success"
        Path     = $agentPath
        SizeMB   = [math]::Round($size,2)
        Time     = $ts
    }

    # Install agent inline
    Write-Host "`n🚀 Installing agent..." -ForegroundColor Cyan
    Write-Host "& $agentPath -c $company -e $tenant -j $key -i" -ForegroundColor Yellow
    try {
        & $agentPath -c $company -e $tenant -j $key -i
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Completed"; Time = (Get-Date) }
    } catch {
        Write-Host "❌ Install failed: $_" -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Failed: $_"; Time = (Get-Date) }
    }

    Export-Data -Object $log -BaseName $baseName
    Write-Host "`n📄 Transcript saved to: $txtPath"

    if (Test-Path $agentPath) {
        Remove-Item $agentPath -Force -ErrorAction SilentlyContinue
        Write-Host "`n🧹 Cleaned up installer from: $agentPath" -ForegroundColor DarkGray
    }

    Pause-Script "Install complete. Use the main menu to zip/email results if needed."
    Stop-Transcript
}

Run-AgentInstall
