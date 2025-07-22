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
    $txtPath = "$env:TEMP\\$baseName.txt"
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
    $agentPath = "$env:TEMP\\cybercnsagent.exe"
    $metaUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    try {
        Write-Host "`n🔎 Resolving agent download URL..." -ForegroundColor Cyan
        $response = Invoke-RestMethod -Uri $metaUrl -UseBasicParsing
        $agentUrl = $response.agentDownloadURL
        Write-Host "⬇️ Downloading from: $agentUrl" -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $agentUrl -OutFile $agentPath -UseBasicParsing
        Write-Host "✅ Agent downloaded to $agentPath" -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Download Agent"; Status = "Success"; Path = $agentPath; Time = $ts }
    } catch {
        Write-Host "❌ Failed to resolve/download agent: $_" -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Download Agent"; Status = "Failed"; Time = $ts }
        Stop-Transcript
        return
    }

    # Run agent installer
    Write-Host "`n🚀 Installing agent..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $agentPath -ArgumentList "-c $company -e $tenant -j $key -i" -Wait
        Write-Host "✅ Agent installed successfully." -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Completed"; Time = $ts }
    } catch {
        Write-Host "❌ Agent install failed: $_" -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Failed"; Time = $ts }
    }

    Export-Data -Object $log -BaseName $baseName
    Write-Host "`n📄 Transcript saved to: $txtPath"

    Run-ZipAndEmailResults
    Run-CleanupExportFolder

    Pause-Script "Install routine complete. Press any key to close."
    Stop-Transcript
}

Run-AgentInstall
