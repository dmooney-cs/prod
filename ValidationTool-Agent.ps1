
# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Tool                       ║
# ╚════════════════════════════════════════════════════╝

function Install-CyberCNSAgent {
    Write-Host "▶ Installing CyberCNS Agent..." -ForegroundColor Cyan
    $company = Read-Host "Enter Company ID"
    $tenant = Read-Host "Enter Tenant ID"
    $key = Read-Host "Enter Secret Key"
    $downloadUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    $tempDir = "C:\Script-Temp"
    $agentPath = "$tempDir\cybercnsagent.exe"

    if (-not (Test-Path $tempDir)) {
        New-Item -Path $tempDir -ItemType Directory | Out-Null
    }

    try {
        $url = Invoke-RestMethod -Uri $downloadUrl -UseBasicParsing
        Invoke-WebRequest -Uri $url -OutFile $agentPath
        Write-Host "Downloaded to: $agentPath"
    } catch {
        Write-Host "❌ Download failed: $_" -ForegroundColor Red
        return
    }

    $cmd = "`"$agentPath`" -c $company -e $tenant -j $key -i"
    Write-Host "Executing: $cmd" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    Invoke-Expression $cmd
    Write-Host "✔ Installation complete." -ForegroundColor Green
}

function Check-AgentStatus {
    Write-Host "▶ Checking Agent Status..." -ForegroundColor Cyan
    $agentService = Get-Service -Name "cybercnsagent" -ErrorAction SilentlyContinue
    if ($agentService -and $agentService.Status -eq "Running") {
        Write-Host "✔ Agent service is running." -ForegroundColor Green
    } else {
        Write-Host "❌ Agent service not running." -ForegroundColor Red
    }

    $logPath = "C:\Program Files (x86)\CyberCNSAgent\logs\cybercns.txt"
    $ver = ""
    if (Test-Path $logPath) {
        $ver = Select-String -Path $logPath -Pattern "Agent Version" | ForEach-Object {
            ($_ -split ":")[1].Trim()
        }
    }
    Write-Host "Agent Version: $ver"

    $checkinJson = "C:\Program Files (x86)\CyberCNSAgent\lastcheckin.json"
    if (Test-Path $checkinJson) {
        $json = Get-Content $checkinJson | ConvertFrom-Json
        $lastCheckin = $json.last_checkin_time
        Write-Host "Last Check-in: $lastCheckin"
    } else {
        Write-Host "Check-in data not found."
    }
}

function Clear-PendingJobs {
    Write-Host "🧹 Clearing Agent Job Queue..." -ForegroundColor Red
    $queuePath = "C:\Program Files (x86)\CyberCNSAgent\pendingjobqueue"
    if (Test-Path $queuePath) {
        Remove-Item "$queuePath\*" -Force -Recurse -ErrorAction SilentlyContinue
        Write-Host "✔ Job queue cleared." -ForegroundColor Green
    } else {
        Write-Host "❌ Job queue path not found."
    }
}

function Enable-SMBForAgent {
    Write-Host "▶ Enabling SMB for Agent Compatibility..." -ForegroundColor Cyan
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name SMB2 -Value 1 -Force
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (SMB-In)" -Enabled True -Profile Any
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (NB-Session-In)" -Enabled True -Profile Any
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name LocalAccountTokenFilterPolicy -PropertyType DWord -Value 1 -Force
    Write-Host "✔ SMB enabled and firewall adjusted." -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportDir\AgentExport_$hostname_$timestamp.zip"

    if (Test-Path $exportDir) {
        Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile -Force
        Write-Host "✔ ZIP file created: $zipFile" -ForegroundColor Green

        $to = Read-Host "Enter recipient email"
        $mailto = "mailto:$to?subject=Agent%20Export%20Results&body=Results%20attached.%20ZIP:%20$zipFile"
        Start-Process $mailto
    } else {
        Write-Host "❌ No export folder found." -ForegroundColor Yellow
    }
}

function Run-CleanupScriptData {
    Write-Host "🧹 Cleaning up export folder..." -ForegroundColor Red
    Remove-Item -Path "C:\Script-Export\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "✔ Cleanup complete."
}

function Show-AgentMenu {
    do {
        Write-Host "`n========= 🛠 Agent Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Install Agent"
        Write-Host "[2] Check Agent Status"
        Write-Host "[3] Clear Agent Job Queue"
        Write-Host "[4] Enable SMB for Agent"
        Write-Host "[5] Zip and Email Results"
        Write-Host "[6] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Install-CyberCNSAgent }
            "2" { Check-AgentStatus }
            "3" { Clear-PendingJobs }
            "4" { Enable-SMBForAgent }
            "5" { Run-ZipAndEmailResults }
            "6" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-AgentMenu
