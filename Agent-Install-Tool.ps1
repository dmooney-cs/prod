# ==========================================
#   ConnectSecure Agent Install Utility
# ==========================================

. ([scriptblock]::Create((irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 -UseBasicParsing)))
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
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Uninstall-CyberCNSAgentV4.ps1 | iex
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

    # Re-import functions in case transcript or host wiped scope
    . ([scriptblock]::Create((irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 -UseBasicParsing)))

    Run-ZipAndEmailResults
    Run-CleanupExportFolder

    if (Test-Path $agentPath) {
        Remove-Item $agentPath -Force -ErrorAction SilentlyContinue
        Write-Host "`n🧹 Cleaned up installer from: $agentPath" -ForegroundColor DarkGray
    }

    Pause-Script "Install routine complete. Press any key to close."
    Stop-Transcript
}

Run-AgentInstall
