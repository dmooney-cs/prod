
function Run-BrowserExtensionDetails {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    $outputCsv = "$outputDir\OSquery-browserext-$timestamp-$hostname.csv"
    $outputJson = "$outputDir\RawBrowserExt-$timestamp-$hostname.json"

    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory | Out-Null
    }

    $osqueryPaths = @(
        "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe",
        "C:\Windows\CyberCNSAgent\osqueryi.exe"
    )

    $osquery = $osqueryPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $osquery) {
        Write-Host "❌ osqueryi.exe not found in expected locations." -ForegroundColor Red
        Read-Host -Prompt "Press any key to exit"
        return
    }

    Push-Location (Split-Path $osquery)

    $sqlQuery = @"
SELECT name, browser_type, version, path, sha1(name || path) AS unique_id 
FROM chrome_extensions 
WHERE chrome_extensions.uid IN (SELECT uid FROM users) 
GROUP BY unique_id;
"@

    $json = & .\osqueryi.exe --json "$sqlQuery" 2>$null
    Pop-Location

    Set-Content -Path $outputJson -Value $json

    if (-not $json) {
        Write-Host "⚠️ No browser extension data returned by osquery." -ForegroundColor Yellow
        Write-Host "`nRaw JSON saved for review:" -ForegroundColor DarkGray
        Write-Host "$outputJson" -ForegroundColor Cyan
        Read-Host -Prompt "Press any key to continue"
        return
    }

    try {
        $parsed = $json | ConvertFrom-Json

        if ($parsed.Count -eq 0) {
            Write-Host "⚠️ No browser extensions found." -ForegroundColor Yellow
        } else {
            Write-Host "`n✅ Extensions found: $($parsed.Count)" -ForegroundColor Green
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            foreach ($ext in $parsed) {
                Write-Host "• $($ext.name) ($($ext.browser_type)) — $($ext.version)" -ForegroundColor White
            }
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            $parsed | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
        }

    } catch {
        Write-Host "❌ Failed to parse or export data." -ForegroundColor Red
        Read-Host -Prompt "Press any key to continue"
        return
    }

    Write-Host "`n📁 Output Files:" -ForegroundColor Yellow
    Write-Host "• CSV:  $outputCsv" -ForegroundColor Cyan
    Write-Host "• JSON: $outputJson" -ForegroundColor Cyan
    Read-Host -Prompt "`nPress any key to exit"
}



# Data-Collection-and-Validation-Tool.ps1
# ConnectSecure - System Collection and Validation Launcher

Write-Host "`n======== Data Collection and Validation Tool ========" -ForegroundColor Cyan

function Show-MainMenu {
    Write-Host ""
    Write-Host "1. Validation Scripts"
    Write-Host "   → Run application, driver, network, and update validations."
    Write-Host ""
    Write-Host "2. Agent Maintenance"
    Write-Host "   → Install, update, or troubleshoot the ConnectSecure agent."
    Write-Host ""
    Write-Host "3. Probe Troubleshooting"
    Write-Host "   → Diagnose probe issues and test scanning tools."
    Write-Host ""
    Write-Host "4. Zip and Email Results"
    Write-Host "   → Package collected data into a ZIP and open your mail client."
    Write-Host ""
    Write-Host "Q. Close and Purge Script Data"
    Write-Host "   → Optionally email, then delete all script-related files."
    Write-Host ""
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "`n[Simulated] Agent Maintenance..." -ForegroundColor Cyan; Pause }
        "3" { Write-Host "`n[Simulated] Probe Troubleshooting..." -ForegroundColor Cyan; Pause }
        "4" { Write-Host "`n[Simulated] Zipping results and launching email..." -ForegroundColor Cyan; Pause }
        "Q" { Purge-ScriptData }
        default { Write-Host "`nInvalid option. Please select again." -ForegroundColor Red }
    }
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Browser Extension details"
        Write-Host "4. Back to Main Menu"
        $subChoice = Read-Host "Select an option"
        switch ($subChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-BrowserExtensionDetails }
            "4" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-OfficeValidation {
    Write-Host "Scanning for installed applications..." -ForegroundColor Cyan
    $appFilter = Read-Host "Enter a keyword to filter applications (or press Enter to list all)"
    $results = @()

    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($regPath in $registryPaths) {
        Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DisplayName) {
                if ($appFilter -eq "" -or $_.DisplayName -like "*$appFilter*") {
                    $category = "Application"
                    $registryPath = $_.PSPath -replace "^Microsoft.PowerShell.Core\Registry::", ""
                    $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }

                    $results += [PSCustomObject]@{
                        Name             = $_.DisplayName
                        Version          = $_.DisplayVersion
                        Publisher        = $_.Publisher
                        InstallLocation  = $installLocation
                        InstallDate      = $_.InstallDate
                        InstalledBy      = $_.InstallSource
                        Source           = "Registry"
                        Category         = $category
                        RegistryPath     = $registryPath
                        ProfileType      = "N/A"
                    }
                }
            }
        }
    }

    Get-AppxPackage | ForEach-Object {
        if ($appFilter -eq "" -or $_.Name -like "*$appFilter*" -or $_.PackageFullName -like "*$appFilter*") {
            $category = "Microsoft Store App"
            $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }

            $results += [PSCustomObject]@{
                Name             = $_.Name
                Version          = $_.Version
                Publisher        = $_.Publisher
                InstallLocation  = $installLocation
                InstallDate      = "N/A"
                InstalledBy      = "N/A"
                Source           = "Microsoft Store"
                Category         = $category
                RegistryPath     = "N/A"
                ProfileType      = "N/A"
            }
        }
    }

    $teamsPaths = Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Teams\Teams.exe" -Recurse -ErrorAction SilentlyContinue
    foreach ($teamsPath in $teamsPaths) {
        $version = (Get-Item $teamsPath.FullName).VersionInfo.FileVersion
        $installLocation = $teamsPath.DirectoryName

        $results += [PSCustomObject]@{
            Name             = 'Microsoft Teams'
            Version          = $version
            Publisher        = 'Microsoft'
            InstallLocation  = $installLocation
            InstallDate      = "N/A"
            InstalledBy      = "N/A"
            Source           = "Teams (Local)"
            Category         = "Microsoft Store App"
            RegistryPath     = "N/A"
            ProfileType      = "N/A"
        }
    }

    if ($results.Count -eq 0) {
        Write-Host "No matching applications found." -ForegroundColor Yellow
    } else {
        Write-Host "`nDetected Applications:" -ForegroundColor Green
        $results | Sort-Object Name | Format-Table Name, Version, Publisher, InstallLocation, InstallDate, InstalledBy, Category, Source -AutoSize

        $dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $hostname = $env:COMPUTERNAME
        $exportPath = "C:\Script-Export"

        if (-not (Test-Path -Path $exportPath)) {
            New-Item -Path $exportPath -ItemType Directory
            Write-Host "Created folder: $exportPath" -ForegroundColor Cyan
        }

        $csvPath = "$exportPath\All Applications Detected - $dateTime - $hostname.csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nResults exported to: $csvPath" -ForegroundColor Green
    }
}

function Run-DriverValidation {
    $exportFolder = "C:\Script-Export"
    if (-not (Test-Path $exportFolder)) {
        New-Item -Path $exportFolder -ItemType Directory | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputPath = "$exportFolder\Installed_Drivers_${hostname}_$timestamp.csv"

    $drivers = Get-WmiObject Win32_PnPSignedDriver | Sort-Object DeviceName

    $driverInfo = $drivers | Select-Object `
        DeviceName,
        DeviceID,
        DriverVersion,
        DriverProviderName,
        InfName,
        @{Name="DriverDate";Expression={if ($_.DriverDate) { 
            [datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null) 
        } else { $null }}},
        Manufacturer,
        DriverPath,
        @{Name="BestGuessDriverFullPath";Expression={
            $sysFile = $_.DriverPath
            if ($sysFile -and $sysFile.EndsWith(".sys")) {
                $fileName = [System.IO.Path]::GetFileName($sysFile)
                $paths = @(
                    "C:\Windows\System32\drivers\$fileName",
                    "C:\Windows\System32\DriverStore\FileRepository\$fileName"
                )
                foreach ($p in $paths) {
                    if (Test-Path $p) { return $p }
                }
                return $paths[0]
            } else {
                return "Unknown"
            }
        }},
        @{Name="FullInfFilePath";Expression={
            $infFile = $_.InfName
            $infPaths = @(
                "C:\Windows\INF\$infFile",
                "C:\Windows\System32\DriverStore\FileRepository\$infFile"
            )
            foreach ($p in $infPaths) {
                if (Test-Path $p) { return $p }
            }
            return "INF not found"
        }}

    $driverInfo | Format-Table -AutoSize
    $driverInfo | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nInstalled driver summary exported to:" -ForegroundColor Cyan
    Write-Host $outputPath -ForegroundColor Green

    Read-Host -Prompt "Press Enter to return"
}

function Purge-ScriptData {
    $tempPath = "C:\Script-Temp"
    if (Test-Path $tempPath) {
        $files = Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        $count = 0
        foreach ($item in $files) {
            try {
                Remove-Item -Path $item.FullName -Force -Recurse -ErrorAction Stop
                Write-Host "Deleted: $($item.FullName)" -ForegroundColor DarkGray
                $count++
            } catch {
                Write-Host "Failed to delete: $($item.FullName)" -ForegroundColor Red
            }
        }
        Write-Host "`nTotal items deleted: $count" -ForegroundColor Cyan
    } else {
        Write-Host "No temp folder found at $tempPath." -ForegroundColor Yellow
    }
    Write-Host "`nPress Enter to exit..."
    Read-Host | Out-Null
    exit
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
        if ($choice -ne "Q") {
            Write-Host "`nPress Enter to return to menu..."
            Read-Host | Out-Null
        }
        Clear-Host
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
