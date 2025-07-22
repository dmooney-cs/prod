# ==========================================
#   ConnectSecure Agent Install Utility
# ==========================================

. { iwr -useb "https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1" } | iex
Show-Header "ConnectSecure Agent Install Utility"

function Run-AgentInstall {
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $log = @()
    $folder = Ensure-ExportFolder
    $csvPath = "$folder\AgentInstallLog_$ts`_$hn.csv"
    $txtPath = "$folder\AgentInstallLog_$ts`_$hn.txt"
    Start-Transcript -Path $txtPath -Append

    # Check for running services
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

    # Prompt for install parameters
    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $key     = Read-Host "Enter Secret Key"

    # Download agent using resolved URL
    $agentMetaUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    $agentPath = "$env:TEMP\cybercnsagent.exe"
    try {
        Write-Host "`n🔎 Resolving agent download URL..." -ForegroundColor Cyan
        $agentMeta = Invoke-RestMethod -Uri $agentMetaUrl -UseBasicParsing
        $agentUrl = $agentMeta.agentDownloadURL
        Write-Host "⬇️ Downloading from: $agentUrl" -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $agentUrl -OutFile $agentPath -UseBasicParsing
        Write-Host "✅ Agent downloaded to $agentPath" -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Download Agent"; Status = "Success"; Path = $agentPath; Time = $ts }
    } catch {
        Write-Host "❌ Failed to resolve/download agent: $_" -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Download Agent"; Status = "Failed: $_"; Time = $ts }
        return
    }

    # Install agent
    Write-Host "`n🚀 Installing agent..." -ForegroundColor Cyan
    $installCmd = "& `"$agentPath`" -c $company -e $tenant -j $key -i"
    Write-Host "Running: $installCmd" -ForegroundColor Yellow
    try {
        Start-Process -FilePath $agentPath -ArgumentList "-c $company -e $tenant -j $key -i" -Wait
        Write-Host "✅ Agent installed successfully." -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Completed"; Time = $ts }
    } catch {
        Write-Host "❌ Agent install failed: $_" -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Install Agent"; Status = "Failed: $_"; Time = $ts }
    }

    Export-Data -Data $log -Path $csvPath
    Write-Host "`n📁 CSV log exported to: " -NoNewline; Write-ExportPath $csvPath
    Write-Host "📄 Transcript saved to: " -NoNewline; Write-ExportPath $txtPath

    Run-ZipAndEmailResults
    Run-CleanupExportFolder

    Pause-Script "Install routine complete. Press any key to close."
    Stop-Transcript
}

Run-AgentInstall
