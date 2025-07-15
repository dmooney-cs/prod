<#
.SYNOPSIS
    CyberCNS Agent Maintenance Script

.DESCRIPTION
    - Displays service status and starts if needed
    - Prompts before uninstall if services fail
    - Uses uninstall.bat or raw batch fallback
    - Downloads installer via API securely
    - Substitutes Company, Tenant, and Key into final install command
    - Logs to C:\Script-Export and returns to main menu on completion
#>

# === Setup ===
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$shortDate = Get-Date -Format "yyyy-MM-dd"
$shortTime = Get-Date -Format "HHmm"
$hostname = $env:COMPUTERNAME
$exportDir = "C:\Script-Export"
if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
$logFile = "$exportDir\$hostname-AgentMaintenaceScriptLog-$shortDate-$shortTime.txt"
Start-Transcript -Path $logFile -Append

# === Constants ===
$agentDir = "C:\Program Files (x86)\CyberCNSAgent"
$uninstallBatPath = Join-Path $agentDir "uninstall.bat"
$rawUninstall = @"
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

# === Function: Run Installer ===
function Run-Install {
    $companyId = Read-Host "Enter Company ID"
    $tenantId  = Read-Host "Enter Tenant ID"
    $secretKey = Read-Host "Enter Secret Key"

    Write-Host "`nUsing TLS 1.2 for secure agent link download..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Host "Fetching agent download URL from ConnectSecure API..." -ForegroundColor Cyan
    try {
        $source = Invoke-RestMethod -Method "Get" -Uri "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    } catch {
        Write-Host "Failed to retrieve download URL: $_" -ForegroundColor Red
        return
    }

    $downloadDir = "C:\Script-Temp"
    if (-not (Test-Path $downloadDir)) {
        New-Item -Path $downloadDir -ItemType Directory | Out-Null
    }

    $destination = Join-Path $downloadDir "cybercnsagent.exe"

    Write-Host "Downloading agent to $destination" -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $source -OutFile $destination -UseBasicParsing
        Write-Host "Agent downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download agent: $_" -ForegroundColor Red
        return
    }

    $installCmd = "$destination -c $companyId -e $tenantId -j $secretKey -i"
    Write-Host "`nExecuting: $installCmd" -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    cmd /c $installCmd

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Agent installation completed successfully." -ForegroundColor Green
    } else {
        Write-Host "Agent installation failed (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
    }

    # Final prompt to exit after installation
    Read-Host -Prompt "`nPress any key to exit"
}

# === Function: Run Uninstall Sequence ===
function Run-Uninstall {
    Write-Host "`nStarting uninstall process..." -ForegroundColor Yellow

    if (Test-Path $uninstallBatPath) {
        $uninstallScript = Get-Content $uninstallBatPath -Raw
        Write-Host "Loaded uninstall.bat from $hostname." -ForegroundColor Green
    } else {
        $uninstallScript = $rawUninstall
        Write-Host "Using built-in uninstall routine (raw fallback)." -ForegroundColor Yellow
    }

    if (Test-Path "$agentDir\cybercnsagent.exe") {
        Write-Host "`nRunning initial uninstall trigger..." -ForegroundColor Cyan
        try {
            & "$agentDir\cybercnsagent.exe" -r
        } catch {
            Write-Host "Error running cybercnsagent.exe -r: $_" -ForegroundColor Red
        }
    }

    $tempBat = "$env:TEMP\_agent_uninstall.bat"
    $uninstallScript | Out-File -FilePath $tempBat -Encoding ASCII -Force
    Write-Host "`nExecuting uninstall script..." -ForegroundColor Cyan
    cmd /c $tempBat
    Remove-Item $tempBat -Force -ErrorAction SilentlyContinue
    Write-Host "Uninstall process completed." -ForegroundColor Green

    # Prompt to install agent after uninstall
    Read-Host -Prompt "`nPress any key to install agent"
}

# === Function: Check + Start Services ===
function Test-Services {
    $svc1 = Get-Service -Name "CyberCNSAgent" -ErrorAction SilentlyContinue
    $svc2 = Get-Service -Name "ConnectSecureAgentMonitor" -ErrorAction SilentlyContinue
    $runningCount = 0

    Write-Host "`n=== Agent Service Status ===" -ForegroundColor Cyan

    if ($svc1) {
        if ($svc1.Status -ne "Running") {
            try {
                Start-Service $svc1.Name -ErrorAction Stop
                Write-Host "Started CyberCNSAgent" -ForegroundColor Green
            } catch {
                Write-Host "Could not start CyberCNSAgent" -ForegroundColor Red
            }
        }
        $svc1.Refresh()
        if ($svc1.Status -eq "Running") { $runningCount++ }
        $color1 = if ($svc1.Status -eq "Running") { "Green" } else { "Red" }
        Write-Host "CyberCNSAgent: $($svc1.Status)" -ForegroundColor $color1
    } else {
        Write-Host "CyberCNSAgent: Not Found" -ForegroundColor Yellow
    }

    if ($svc2) {
        if ($svc2.Status -ne "Running") {
            try {
                Start-Service $svc2.Name -ErrorAction Stop
                Write-Host "Started ConnectSecureAgentMonitor" -ForegroundColor Green
            } catch {
                Write-Host "Could not start ConnectSecureAgentMonitor" -ForegroundColor Red
            }
        }
        $svc2.Refresh()
        if ($svc2.Status -eq "Running") { $runningCount++ }
        $color2 = if ($svc2.Status -eq "Running") { "Green" } else { "Red" }
        Write-Host "ConnectSecureAgentMonitor: $($svc2.Status)" -ForegroundColor $color2
    } else {
        Write-Host "ConnectSecureAgentMonitor: Not Found" -ForegroundColor Yellow
    }

    return $runningCount
}

# === MAIN LOGIC ===
$running = Test-Services

if ($running -eq 2) {
    do {
        Write-Host "`n========= ACTION MENU =========" -ForegroundColor Cyan
        Write-Host "1. Install Agent"
        Write-Host "2. Uninstall Agent"
        Write-Host "3. Exit Script"
        $choice = Read-Host "`nSelect an option (1-3)"

        switch ($choice) {
            "1" { Run-Install }
            "2" { Run-Uninstall }
            "3" {
                Write-Host "`nExiting script..." -ForegroundColor Yellow
                break
            }
            default {
                Write-Host "Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red
            }
        }

        Write-Host ""
        Read-Host -Prompt "Press any key to return to the main menu"

    } while ($true)
} else {
    Write-Host "`nOne or more services failed to start or are not running." -ForegroundColor Red
    $prompt = Read-Host "Would you like to uninstall and reinstall the agent now? (Y/N)"
    if ($prompt -match '^[Yy]$') {
        Run-Uninstall
        Run-Install
    } else {
        Write-Host "Uninstall cancelled. No actions were taken." -ForegroundColor Yellow
    }
}

Stop-Transcript
Write-Host "`n=== Script execution completed ===" -ForegroundColor Green
Write-Host "Log file saved to: $logFile" -ForegroundColor Yellow
