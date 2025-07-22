# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Maintenance Menu                ║
# ║ Version: 1.0 | 2025-07-21                                  ║
# ║ Includes: Agent Status, SMB Tools, Clear Jobs             ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-CheckAgentStatus {
    Show-Header "Check CyberCNS Agent Status"

    $service = Get-Service -Name "cybercnsagent" -ErrorAction SilentlyContinue
    $version = ""
    $lastSeen = ""
    $status = "Unknown"
    $host = $env:COMPUTERNAME
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    if ($service) {
        $status = $service.Status
        $log = "C:\Program Files (x86)\CyberCNSAgent\logs\cybercns.txt"
        if (Test-Path $log) {
            $version = Select-String -Path $log -Pattern "Agent Version\s+:" | Select-Object -First 1 | ForEach-Object { $_.Line.Split(":")[1].Trim() }
        }
        $json = "C:\Program Files (x86)\CyberCNSAgent\lastcheckin.json"
        if (Test-Path $json) {
            $data = Get-Content $json | ConvertFrom-Json
            $lastSeen = Get-Date ($data.lastCheckin) -Format "yyyy-MM-dd HH:mm:ss"
        }
    }

    $info = [PSCustomObject]@{
        Hostname     = $host
        Status       = $status
        Version      = $version
        LastCheckIn  = $lastSeen
        CheckedAt    = $timestamp
    }

    Export-Data -Object $info -BaseName "AgentStatus"
    Pause-Script
}

function Run-ClearPendingJobs {
    Show-Header "Clear Pending Job Queue"
    $path = "C:\Program Files (x86)\CyberCNSAgent\pendingjobqueue"
    $count = 0

    if (Test-Path $path) {
        try {
            $files = Get-ChildItem $path -File
            $count = $files.Count
            $files | Remove-Item -Force
            Write-Host "Cleared $count pending job(s)." -ForegroundColor Green
        } catch {
            Write-Host "Error clearing job queue: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Job queue not found." -ForegroundColor Yellow
    }

    Pause-Script
}

function Run-SetSMB {
    Show-Header "Set SMB and Enable Firewall Rules"

    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name SMB2 -Type DWord -Value 1 -Force
        Set-NetFirewallRule -DisplayName "File And Printer Sharing (SMB-In)" -Enabled True -Profile Any
        Set-NetFirewallRule -DisplayName "File And Printer Sharing (NB-Session-In)" -Enabled True -Profile Any
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name LocalAccountTokenFilterPolicy -Type DWord -Value 1 -Force
        Write-Host "✅ SMB 2 enabled and firewall rules configured." -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to set SMB settings." -ForegroundColor Red
    }

    Pause-Script
}

function Run-CheckSMB {
    Show-Header "Check SMB Version Status"
    $output = @()
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"

    $smb1 = Get-ItemProperty -Path $regPath -Name SMB1 -ErrorAction SilentlyContinue
    $smb2 = Get-ItemProperty -Path $regPath -Name SMB2 -ErrorAction SilentlyContinue

    $output += [PSCustomObject]@{
        SMBVersion = "SMB1"
        Enabled    = if ($smb1.SMB1 -eq 1) { "Yes" } else { "No" }
        DisableKey = "Set-ItemProperty -Path `"$regPath`" -Name SMB1 -Value 0"
    }

    $output += [PSCustomObject]@{
        SMBVersion = "SMB2"
        Enabled    = if ($smb2.SMB2 -eq 1) { "Yes" } else { "No" }
        DisableKey = "Set-ItemProperty -Path `"$regPath`" -Name SMB2 -Value 0"
    }

    Export-Data -Object $output -BaseName "SMB_Version_Status"
    Pause-Script
}

function Show-AgentMaintenanceMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Agent Maintenance     ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Check Agent Status",
        " [2] Clear Pending Jobs",
        " [3] Set SMB Configuration",
        " [4] Check SMB Version",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-CheckAgentStatus }
        "2" { Run-ClearPendingJobs }
        "3" { Run-SetSMB }
        "4" { Run-CheckSMB }
        "Q" { return }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Pause-Script
        }
    }
    Show-AgentMaintenanceMenu
}

Show-AgentMaintenanceMenu
