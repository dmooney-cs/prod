# ╔════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool A                     ║
# ║ Version: A.7 – Dual-method Roaming Profile App Detection   ║
# ╚════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-OfficeValidation {
    Clear-Host
    Write-Host "`n=== Microsoft Office Installations ===`n" -ForegroundColor Cyan
    $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                              HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match "Office|Microsoft 365|Word|Excel|Outlook" } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
    Export-Data -Object $apps -BaseName "OfficeAudit"
    Pause-Script
}

function Run-DriverAudit {
    Clear-Host
    Write-Host "`n=== Installed Driver Summary ===`n" -ForegroundColor Cyan
    $drivers = Get-WmiObject Win32_PnPSignedDriver | 
               Select-Object DeviceName, DriverVersion, DriverProviderName, DriverDate
    Export-Data -Object $drivers -BaseName "DriverAudit"
    Pause-Script
}

function Run-RoamingProfileApps {
    Clear-Host
    Write-Host "`n=== Roaming Profile Applications (Dual Method) ===`n" -ForegroundColor Cyan
    $results = @()
    $profiles = Get-WmiObject Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath -like "C:\Users\*" }

    foreach ($profile in $profiles) {
        $username = ($profile.LocalPath -split '\\')[-1]
        $sid = $profile.SID
        $profileType = if ($profile.RoamingConfigured) { "Roaming" } elseif ($profile.Loaded) { "Active" } else { "Local/Old" }
        $keyPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Uninstall\"

        if (Test-Path $keyPath) {
            Get-ChildItem $keyPath -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    if ($app.DisplayName) {
                        $results += [PSCustomObject]@{
                            SID         = $sid
                            User        = $username
                            ProfileType = $profileType
                            AppName     = $app.DisplayName
                            Version     = $app.DisplayVersion
                        }
                    }
                } catch {}
            }
        }
    }

    Export-Data -Object $results -BaseName "RoamingApps"
    Pause-Script
}

function Run-BrowserExtensionDetails {
    Clear-Host
    Write-Host "`n=== Browser Extensions Audit ===`n" -ForegroundColor Cyan
    $paths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data",
        "$env:APPDATA\Mozilla\Firefox\Profiles"
    )
    $results = @()
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $browser = if ($path -match "Chrome") { "Chrome" }
                       elseif ($path -match "Edge") { "Edge" }
                       elseif ($path -match "Firefox") { "Firefox" }
            Get-ChildItem $path -Recurse -Include manifest.json -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $json = Get-Content $_.FullName -Raw | ConvertFrom-Json
                    $results += [PSCustomObject]@{
                        Browser      = $browser
                        Extension    = $json.name
                        Version      = $json.version
                        Description  = $json.description
                        Path         = $_.Directory.FullName
                    }
                } catch {}
            }
        }
    }
    $results = $results | Sort-Object Browser, Extension
    Export-Data -Object $results -BaseName "BrowserExtensions"
    Pause-Script
}

function Show-ValidationMenuA {
    Clear-Host
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        🧪 Validation Tool A – Collection Menu            ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Validate Microsoft Office Installations"
    Write-Host "2. Audit Installed Drivers"
    Write-Host "3. Scan Roaming Profiles for Applications"
    Write-Host "4. Detect Browser Extensions (Chrome, Edge, Firefox)"
    Write-Host ""
    Write-Host "Z. Zip and Email Results"
    Write-Host "C. Cleanup Export Folder"
    Write-Host "Q. Quit to Main Menu"
    Write-Host ""
}

do {
    Show-ValidationMenuA
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Run-OfficeValidation }
        '2' { Run-DriverAudit }
        '3' { Run-RoamingProfileApps }
        '4' { Run-BrowserExtensionDetails }
        'Z' { Run-ZipAndEmailResults }
        'C' { Run-CleanupExportFolder }
        'Q' { return }
        default {
            Write-Host "Invalid selection. Try again." -ForegroundColor Yellow
            Pause-Script
        }
    }
} while ($true)
