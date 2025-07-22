# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool B                      ║
# ║ Version: B.1 | 2025-07-21                                   ║
# ║ Includes VC++ Detection, Windows Patching, ZIP, Cleanup     ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-VCPPValidation {
    Show-Header "VC++ Redistributables + Dependency Audit"
    $results = @()

    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($key in $keys) {
        Get-ChildItem $key | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath
            if ($props.DisplayName -like "*Visual C++*Redistributable*") {
                $results += [PSCustomObject]@{
                    Type    = "Installed Redistributable"
                    Name    = $props.DisplayName
                    Version = $props.DisplayVersion
                }
            }
        }
    }

    Export-Data -Object $results -BaseName "VCPP_Redistributables"
    Pause-Script
}

function Run-WindowsPatchCheck {
    Show-Header "Windows Patch Details – HotFix / WMIC / Osquery"

    # Get-HotFix
    try {
        $hotfix = Get-HotFix | Select-Object Description, HotFixID, InstalledOn
        Export-Data -Object $hotfix -BaseName "Patches_GetHotFix"
    } catch {
        Write-Host "Get-HotFix failed." -ForegroundColor Red
    }

    # WMIC
    try {
        $wmic = wmic qfe list full /format:csv | Out-String
        $outFile = Get-ExportPath -BaseName "Patches_WMIC" -Ext "txt"
        $wmic | Out-File $outFile -Encoding UTF8
        Write-ExportPath $outFile
    } catch {
        Write-Host "WMIC failed." -ForegroundColor Red
    }

    # Osquery (if available)
    $osquery = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    if (Test-Path $osquery) {
        $query = "SELECT hotfix_id, description, installed_on FROM patches;"
        try {
            $raw = & "$osquery" --json "$query"
            $parsed = $raw | ConvertFrom-Json
            Export-Data -Object $parsed -BaseName "Patches_Osquery" -Ext "json"
            Export-Data -Object $parsed -BaseName "Patches_Osquery" -Ext "csv"
        } catch {
            Write-Host "Osquery failed to run." -ForegroundColor Red
        }
    } else {
        Write-Host "osqueryi.exe not found." -ForegroundColor Yellow
    }

    Pause-Script
}

function Show-CollectionMenuB {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Collection Tool B     ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] VC++ Redistributables + Dependency Check",
        " [2] Windows Patch Validation (HotFix, WMIC, Osquery)",
        " [3] Zip and Email Results",
        " [4] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-VCPPValidation }
        "2" { Run-WindowsPatchCheck }
        "3" { Invoke-ZipAndEmailResults }
        "4" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-CollectionMenuB
}

Show-CollectionMenuB
