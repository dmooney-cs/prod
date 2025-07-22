# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – System Info B                         ║
# ║ Version: 1.0 | 2025-07-21                                  ║
# ║ Checks: Reboot Status, Logs, Startup, ZIP, Cleanup         ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-PendingRebootCheck {
    Show-Header "Pending Reboot Check"
    $results = @()

    $results += [PSCustomObject]@{
        Check  = "RebootPending (CCM)"
        Result = Test-Path "HKLM:\SOFTWARE\Microsoft\CCM\RebootRequired"
    }
    $results += [PSCustomObject]@{
        Check  = "PendingFileRenameOperations"
        Result = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue).Count -gt 0
    }
    $results += [PSCustomObject]@{
        Check  = "Windows Update RebootRequired"
        Result = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    }
    $results += [PSCustomObject]@{
        Check  = "Component Based Servicing"
        Result = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
    }

    Export-Data -Object $results -BaseName "PendingRebootStatus"
    Pause-Script
}

function Run-EventLogSummary {
    Show-Header "Event Log Summary (72 hours)"
    $cutoff = (Get-Date).AddDays(-3)
    $logs = @("System", "Application")
    $entries = @()

    foreach ($log in $logs) {
        try {
            $errors = Get-WinEvent -FilterHashtable @{LogName=$log; Level=2; StartTime=$cutoff}
            $entries += $errors | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
        } catch {
            $entries += [PSCustomObject]@{
                TimeCreated = ""
                ProviderName = $log
                Id = "N/A"
                LevelDisplayName = "Error"
                Message = "Could not read log"
            }
        }
    }

    Export-Data -Object $entries -BaseName "EventLogErrors"
    Pause-Script
}

function Run-StartupAudit {
    Show-Header "Startup / Autostart Item Audit"
    $results = @()

    $results += Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -ErrorAction SilentlyContinue |
        Get-Member -MemberType NoteProperty | ForEach-Object {
            [PSCustomObject]@{ Source = "HKLM Run"; Name = $_.Name }
        }

    $results += Get-ScheduledTask | Where-Object { $_.State -ne 'Disabled' } | ForEach-Object {
        [PSCustomObject]@{ Source = "Scheduled Task"; Name = $_.TaskName }
    }

    $results += Get-Service | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' } | ForEach-Object {
        [PSCustomObject]@{ Source = "Service (Auto)"; Name = $_.Name }
    }

    Export-Data -Object $results -BaseName "StartupItems"
    Pause-Script
}

function Show-SystemInfoBMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – System Info B         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Pending Reboot Check",
        " [2] Event Log Error Summary",
        " [3] Startup / Autostart Audit",
        " [4] Zip and Email Results",
        " [5] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-PendingRebootCheck }
        "2" { Run-EventLogSummary }
        "3" { Run-StartupAudit }
        "4" { Invoke-ZipAndEmailResults }
        "5" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-SystemInfoBMenu
}

Show-SystemInfoBMenu
<Recovered System Info B>
