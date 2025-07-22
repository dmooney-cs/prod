# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – System Info A                         ║
# ║ Version: 1.0 | 2025-07-21                                  ║
# ║ Checks: Firewall, Defender, Disk + SMART, ZIP, Cleanup     ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-FirewallDefenderCheck {
    Show-Header "Firewall & Defender Status"
    $fw = Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction
    $def = Get-MpComputerStatus | Select-Object AMServiceEnabled, RealTimeProtectionEnabled, AntivirusEnabled, AntivirusSignatureVersion

    $results = @()
    $fw | ForEach-Object {
        $results += [PSCustomObject]@{
            Category = "Firewall"
            Profile  = $_.Name
            Enabled  = $_.Enabled
            Inbound  = $_.DefaultInboundAction
            Outbound = $_.DefaultOutboundAction
        }
    }

    $results += [PSCustomObject]@{
        Category  = "Defender"
        Profile   = "N/A"
        Enabled   = $def.AMServiceEnabled
        Inbound   = "RealTime: $($def.RealTimeProtectionEnabled)"
        Outbound  = "AV: $($def.AntivirusEnabled)"
        Signature = $def.AntivirusSignatureVersion
    }

    Export-Data -Object $results -BaseName "Firewall_Defender_Status"
    Pause-Script
}

function Run-DiskAndSMARTStatus {
    Show-Header "Disk Space & SMART Health"

    $drives = Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        [PSCustomObject]@{
            Name    = $_.Name
            UsedGB  = "{0:N1}" -f ($_.Used / 1GB)
            FreeGB  = "{0:N1}" -f ($_.Free / 1GB)
            TotalGB = "{0:N1}" -f (($_.Used + $_.Free) / 1GB)
        }
    }

    try {
        $smart = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus |
        ForEach-Object {
            [PSCustomObject]@{
                Drive          = $_.InstanceName
                PredictFailure = $_.PredictFailure
            }
        }
    } catch {
        $smart = [PSCustomObject]@{ Drive = "Unavailable"; PredictFailure = "WMI Error" }
    }

    Export-Data -Object $drives -BaseName "Disk_Space_Usage"
    Export-Data -Object $smart -BaseName "Disk_SMART_Status"
    Pause-Script
}

function Show-SystemInfoAMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – System Info A         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Firewall & Defender Status",
        " [2] Disk Space & SMART Status",
        " [3] Zip and Email Results",
        " [4] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-FirewallDefenderCheck }
        "2" { Run-DiskAndSMARTStatus }
        "3" { Invoke-ZipAndEmailResults }
        "4" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-SystemInfoAMenu
}

Show-SystemInfoAMenu
<Recovered System Info A>
