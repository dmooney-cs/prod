
# Data-Collection-and-Validation-Tool.ps1
# ConnectSecure - System Collection and Validation Launcher (Integrated Office Validation)

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

function Run-ApplicationValidation {
    do {
        Write-Host "`n--- Application Validation ---" -ForegroundColor Cyan
        Write-Host "1. Scan all installed applications"
        Write-Host "2. Scan using a wildcard search term"
        Write-Host "3. Microsoft Office Validation"
        Write-Host "4. Back to Validation Menu"
        $appChoice = Read-Host "Select an option"

        switch ($appChoice) {
            "1" {
                $wildcard = "*"
                Run-OfficeValidation -appFilter $wildcard
            }
            "2" {
                $wildcard = Read-Host "Enter keyword or wildcard to search for (e.g. *Office*)"
                Run-OfficeValidation -appFilter $wildcard
            }
            "3" {
                Run-OfficeValidation -appFilter "*Office*"
            }
            "4" { return }
            default { Write-Host "Invalid option. Returning." -ForegroundColor Red }
        }

        Pause
    } while ($true)
}

function Run-OfficeValidation {
    param([string]$appFilter)

    Write-Host "`nScanning for installed applications matching: $appFilter" -ForegroundColor Cyan
    $results = @()

    $registryPaths = @(
        "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*",
        "HKLM:\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*",
        "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*"
    )

    foreach ($regPath in $registryPaths) {
        Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DisplayName) {
                if ($appFilter -eq "" -or $_.DisplayName.ToLower() -like $appFilter.ToLower()) {
                    $category = "Application"
                    $registryPath = $_.PSPath -replace "^Microsoft.PowerShell.Core\\\\Registry::", ""
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
        if ($appFilter -eq "" -or $_.Name.ToLower() -like $appFilter.ToLower() -or $_.PackageFullName.ToLower() -like $appFilter.ToLower()) {
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

    $teamsPaths = Get-ChildItem "C:\\Users\\*\\AppData\\Local\\Microsoft\\Teams\\Teams.exe" -Recurse -ErrorAction SilentlyContinue
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
        $exportPath = "C:\\Script-Export"

        if (-not (Test-Path -Path $exportPath)) {
            New-Item -Path $exportPath -ItemType Directory
            Write-Host "Created folder: $exportPath" -ForegroundColor Cyan
        }

        $csvPath = "$exportPath\\All Applications Detected - $dateTime - $hostname.csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nResults exported to: $csvPath" -ForegroundColor Green
    }
}

function Run-DriverValidation {
    # Define export path
    $exportFolder = "C:\Script-Export"
    if (-not (Test-Path $exportFolder)) {
        New-Item -Path $exportFolder -ItemType Directory | Out-Null
    }

    # Timestamp and hostname for export file
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputPath = "$exportFolder\Installed_Drivers_${hostname}_$timestamp.csv"

    # Get all installed drivers
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Sort-Object DeviceName

    # Select and enhance driver details
    $driverInfo = $drivers | Select-Object `
        DeviceName,
        DeviceID,
        DriverVersion,
        DriverProviderName,
        InfName,
        @{Name="DriverDate";Expression={if ($_.DriverDate) { 
            [datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null) 
        } else { $null }}}},
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

    # Display results
    $driverInfo | Format-Table -AutoSize

    # Export to CSV
    $driverInfo | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

    # Completion message
    Write-Host "`nInstalled driver summary exported to:" -ForegroundColor Cyan
    Write-Host $outputPath -ForegroundColor Green

    # Pause before exit
    Read-Host -Prompt "Press Enter to exit"
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Application Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Network Validation"
        Write-Host "4. Windows Update Validation"
        Write-Host "5. Back to Main Menu"
        Write-Host "----------------------------------"
        $valChoice = Read-Host "Select an option"

        switch ($valChoice) {
            "1" { Run-ApplicationValidation }
            "2" { Run-DriverValidation }
            "3" { Write-Host "`n[Simulated] Network Validation running..." -ForegroundColor Green; Pause }
            "4" { Write-Host "`n[Simulated] Windows Update Validation running..." -ForegroundColor Green; Pause }
            "5" { return }
            default { Write-Host "Invalid choice. Try again." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-AgentMaintenance { Write-Host "`n[Simulated] Agent Maintenance..." -ForegroundColor Cyan; Pause }
function Run-ProbeTroubleshooting { Write-Host "`n[Simulated] Probe Troubleshooting..." -ForegroundColor Cyan; Pause }
function Run-ZipAndEmail { Write-Host "`n[Simulated] Zipping results and launching email..." -ForegroundColor Cyan; Pause }

function Purge-ScriptData {
    $tempPath = "C:\\Script-Temp"
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

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Run-AgentMaintenance }
        "3" { Run-ProbeTroubleshooting }
        "4" { Run-ZipAndEmail }
        "Q" { Purge-ScriptData }
        default { Write-Host "`nInvalid option. Please select again." -ForegroundColor Red }
    }
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
