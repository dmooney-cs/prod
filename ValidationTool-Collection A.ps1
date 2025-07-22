Ensure-ExportFolder

function Run-OfficeValidation {
    Clear-Host
    $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, `
                              HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match "Office|Microsoft 365|Word|Excel|Outlook" } |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

    if ($apps) {
        irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
        $path = Export-Data -Object $apps -BaseName "OfficeAudit"
        Write-Host "`n📄 Exported file: $path" -ForegroundColor Cyan
    } else {
        Write-Host "No Office applications found." -ForegroundColor Yellow
    }
    Pause-Script
}

function Run-DriverAudit {
    Clear-Host
    $drivers = Get-WmiObject Win32_PnPSignedDriver |
               Select-Object DeviceName, DriverVersion, DriverProviderName, DriverDate

    if ($drivers) {
        irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
        $path = Export-Data -Object $drivers -BaseName "DriverAudit"
        Write-Host "`n📄 Exported file: $path" -ForegroundColor Cyan
    } else {
        Write-Host "No drivers found." -ForegroundColor Yellow
    }
    Pause-Script
}

function Run-RoamingProfileApps {
    Clear-Host
    $results = @()
    $profiles = Get-WmiObject Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath -like "C:\Users\*" }

    foreach ($profile in $profiles) {
        $username = ($profile.LocalPath -split '\\')[-1]
        $sid = $profile.SID
        $type = if ($profile.RoamingConfigured) { "Roaming" } elseif ($profile.Loaded) { "Active" } else { "Local/Old" }
        $keyPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Uninstall\"

        if (Test-Path $keyPath) {
            Get-ChildItem $keyPath -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $app = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    if ($app.DisplayName) {
                        $results += [PSCustomObject]@{
                            SID         = $sid
                            User        = $username
                            ProfileType = $type
                            AppName     = $app.DisplayName
                            Version     = $app.DisplayVersion
                        }
                    }
                } catch {}
            }
        }
    }

    if ($results.Count -gt 0) {
        irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
        $path = Export-Data -Object $results -BaseName "RoamingApps"
        Write-Host "`n📄 Exported file: $path" -ForegroundColor Cyan
    } else {
        Write-Host "No apps found in roaming profiles." -ForegroundColor Yellow
    }
    Pause-Script
}

function Run-BrowserExtensionDetails {
    Clear-Host
    $paths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data",
        "$env:APPDATA\Mozilla\Firefox\Profiles"
    )
    $results = @()
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $browser = if ($path -like "*Chrome*") { "Chrome" }
                       elseif ($path -like "*Edge*") { "Edge" }
                       elseif ($path -like "*Firefox*") { "Firefox" }
            Get-ChildItem $path -Recurse -Include manifest.json -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $json = Get-Content $_.FullName -Raw | ConvertFrom-Json
                    $results += [PSCustomObject]@{
                        Browser     = $browser
                        Extension   = $json.name
                        Version     = $json.version
                        Description = $json.description
                        Path        = $_.Directory.FullName
                    }
                } catch {}
            }
        }
    }

    if ($results.Count -gt 0) {
        irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
        $path = Export-Data -Object $results -BaseName "BrowserExtensions"
        Write-Host "`n📄 Exported file: $path" -ForegroundColor Cyan
    } else {
        Write-Host "No browser extensions found." -ForegroundColor Yellow
    }
    Pause-Script
}

function Run-ZipAndEmailResults {
    irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
    Invoke-ZipAndEmailResults
}

function Run-CleanupExportFolder {
    irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
    Invoke-CleanupExportFolder
}

function Show-ValidationMenuA {
    Clear-Host
    Write-Host ""
    Write-Host "1. Validate Microsoft Office Installations"
    Write-Host "2. Audit Installed Drivers"
    Write-Host "3. Scan Roaming Profiles for Applications"
    Write-Host "4. Detect Browser Extensions (Chrome, Edge, Firefox)"
    Write-Host ""
    Write-Host "Z. Zip and Email Results"
    Write-Host "C. Cleanup Export Folder"
    Write-Host "Q. Quit"
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
            irm https://raw.githubusercontent.com/dmooney-cs/prod/main/Functions-Common.ps1 | iex
            Pause-Script
        }
    }
} while ($true)
