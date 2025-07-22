# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool B                      ║
# ║ Version: B.6 | 2025-07-22                                   ║
# ║ Includes VC++ Detection, Windows Patching, ZIP, Cleanup     ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-VCPPValidation {
    Show-Header "VC++ Redistributables + Dependency Audit"
    $results = @()

    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    foreach ($key in $keys) {
        Get-ChildItem $key -ErrorAction SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            if ($props.DisplayName -like "*Visual C++*Redistributable*") {
                $results += [PSCustomObject]@{
                    Type    = "Installed Redistributable"
                    Name    = $props.DisplayName
                    Version = $props.DisplayVersion
                    Path    = ""
                }
            }
        }
    }

    $folders = @( "$env:ProgramFiles", "$env:ProgramFiles(x86)", "$env:SystemRoot\System32" )
    $vcDlls = @("msvcr", "msvcp", "vcruntime")

    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Include *.dll, *.exe -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $strings = strings $_.FullName | Select-String -Pattern ($vcDlls -join "|")
                    if ($strings) {
                        $results += [PSCustomObject]@{
                            Type    = "Binary Dependency"
                            Name    = ($_.Name)
                            Version = ""
                            Path    = $_.FullName
                        }
                    }
                } catch {}
            }
        }
    }

    if ($results.Count -gt 0) {
        $path = Export-Data -Object $results -BaseName "VCPP_Validation"
        Write-Host "`n📄 Exported file: $path" -ForegroundColor Cyan
    } else {
        Write-Host "No VC++ dependencies or redistributables found." -ForegroundColor Yellow
    }
    Pause-Script
}

function Run-WindowsPatchCheck {
    Show-Header "Windows Patch Details – HotFix / WMIC / Osquery"

    try {
        $hotfix = Get-HotFix | Select-Object Description, HotFixID, InstalledOn
        Export-Data -Object $hotfix -BaseName "Patches_GetHotFix"
    } catch {
        Write-Host "Get-HotFix failed." -ForegroundColor Red
    }

    try {
        $wmic = wmic qfe list full /format:csv | Out-String
        $wmicPath = Get-ExportPath -BaseName "Patches_WMIC" -Ext "txt"
        $wmic | Out-File $wmicPath -Encoding UTF8
        Write-ExportPath $wmicPath
    } catch {
        Write-Host "WMIC failed." -ForegroundColor Red
    }

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

function Run-Zip {
    Invoke-ZipAndEmailResults
}

function Run-Cleanup {
    Invoke-CleanupExportFolder
}

function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Validation Tool B     ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [1] VC++ Redistributables + Dependency Check"
    Write-Host " [2] Windows Patch Validation (HotFix, WMIC, Osquery)"
    Write-Host " [3] Zip and Email Results"
    Write-Host " [4] Cleanup Export Folder"
    Write-Host " [Q] Quit"
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        "1" { Run-VCPPValidation }
        "2" { Run-WindowsPatchCheck }
        "3" { Run-Zip }
        "4" { Run-Cleanup }
        "Q" { return }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Pause-Script
        }
    }
} while ($true)
