# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool A                      ║
# ║ Version: A.3 | Office, Drivers, Dual Roaming, Extensions   ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-OfficeValidation {
    Show-Header "Office Installation Audit"
    $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                              HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match "Office|Microsoft 365|Word|Excel|Outlook" } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
    Export-Data -Object $apps -BaseName "OfficeAudit"
    Pause-Script
}

function Run-DriverAudit {
    Show-Header "Installed Driver Audit"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DriverVersion, Manufacturer, DriverDate
    Export-Data -Object $drivers -BaseName "Installed_Drivers"
    Pause-Script
}

function Run-RoamingProfileApps {
    Show-Header "Roaming Profile Application Audit"
    $results = @()
    $profiles = Get-CimInstance Win32_UserProfile | Where-Object { $_.Special -eq $false }

    foreach ($p in $profiles) {
        $user = $p.LocalPath.Split('\')[-1]

        # Method 1: AppData folders
        $apps = Get-ChildItem "C:\Users\$user\AppData\Local" -Directory -ErrorAction SilentlyContinue
        foreach ($a in $apps) {
            $results += [PSCustomObject]@{
                Profile = $user
                Name    = $a.Name
                Version = ""
                Source  = "AppFolder"
            }
        }

        # Method 2: NTUSER.DAT hive load
        $regPath = "$($p.LocalPath)\NTUSER.DAT"
        if (Test-Path $regPath) {
            try {
                $hive = "HKU\TempHive_$user"
                reg load $hive $regPath | Out-Null
                $entries = Get-ItemProperty "Registry::$hive\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
                foreach ($app in $entries) {
                    if ($app.DisplayName) {
                        $results += [PSCustomObject]@{
                            Profile = $user
                            Name    = $app.DisplayName
                            Version = $app.DisplayVersion
                            Source  = "Registry"
                        }
                    }
                }
                reg unload $hive | Out-Null
            } catch { continue }
        }
    }

    Export-Data -Object $results -BaseName "RoamingProfile_Apps"
    Pause-Script
}

function Run-BrowserExtensionDetails {
    Show-Header "Browser Extension Audit"
    $users = Get-ChildItem "C:\Users" -Force | Where-Object { Test-Path "$($_.FullName)\AppData" }
    $results = @()

    foreach ($user in $users) {
        $base = $user.FullName
        $edge = "$base\AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
        $chrome = "$base\AppData\Local\Google\Chrome\User Data\Default\Extensions"
        $firefox = Get-ChildItem "$base\AppData\Roaming\Mozilla\Firefox\Profiles" -Directory -ErrorAction SilentlyContinue | Select-Object -First 1

        if (Test-Path $chrome) {
            Get-ChildItem $chrome -Directory | ForEach-Object {
                $results += [PSCustomObject]@{ Browser="Chrome"; User=$user.Name; Extension=$_.Name }
            }
        }
        if (Test-Path $edge) {
            Get-ChildItem $edge -Directory | ForEach-Object {
                $results += [PSCustomObject]@{ Browser="Edge"; User=$user.Name; Extension=$_.Name }
            }
        }
        if ($firefox) {
            $extPath = Join-Path $firefox.FullName "extensions"
            if (Test-Path $extPath) {
                Get-ChildItem $extPath -File | ForEach-Object {
                    $results += [PSCustomObject]@{ Browser="Firefox"; User=$user.Name; Extension=$_.Name }
                }
            }
        }
    }

    Export-Data -Object $results -BaseName "Browser_Extensions"
    Pause-Script
}

function Show-CollectionMenuA {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Collection Tool A     ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Microsoft Office Detection",
        " [2] Installed Driver Audit",
        " [3] Roaming Profile Application Scan",
        " [4] Browser Extension Details",
        " [5] Zip and Email Results",
        " [6] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-OfficeValidation }
        "2" { Run-DriverAudit }
        "3" { Run-RoamingProfileApps }
        "4" { Run-BrowserExtensionDetails }
        "5" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-ZipAndEmailResults
        }
        "6" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-CleanupExportFolder
        }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-CollectionMenuA
}

Show-CollectionMenuA
