# Data Collection and Validation Tool
# Compatible with GitHub Raw execution
# Save this as DataValidationTool.ps1 in your repo

function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Microsoft Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Roaming Profile Applications"
        Write-Host "4. Browser Extension Details"
        Write-Host "5. OSQuery Browser Extensions"
        Write-Host "6. SSL Cipher Validation"
        Write-Host "7. Windows Patch Details"
        Write-Host "8. Back to Main Menu"
        $valChoice = Read-Host "Select an option"
        switch ($valChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-BrowserExtensionDetails }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-OfficeValidation {
    $appFilter = Read-Host "Enter a keyword to filter applications (or press Enter to list all)"
    $results = @()
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($path in $regPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DisplayName -and ($appFilter -eq "" -or $_.DisplayName -like "*$appFilter*")) {
                $results += [PSCustomObject]@{
                    Name = $_.DisplayName
                    Version = $_.DisplayVersion
                    Publisher = $_.Publisher
                    InstallLocation = $_.InstallLocation
                    InstallDate = $_.InstallDate
                    Source = "Registry"
                }
            }
        }
    }
    Get-AppxPackage | ForEach-Object {
        if ($appFilter -eq "" -or $_.Name -like "*$appFilter*") {
            $results += [PSCustomObject]@{
                Name = $_.Name
                Version = $_.Version
                Publisher = $_.Publisher
                InstallLocation = $_.InstallLocation
                InstallDate = "N/A"
                Source = "Microsoft Store"
            }
        }
    }
    $teamsPaths = Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Teams\Teams.exe" -Recurse -ErrorAction SilentlyContinue
    foreach ($teamsPath in $teamsPaths) {
        $results += [PSCustomObject]@{
            Name = 'Microsoft Teams'
            Version = (Get-Item $teamsPath.FullName).VersionInfo.FileVersion
            Publisher = 'Microsoft'
            InstallLocation = $teamsPath.DirectoryName
            InstallDate = "N/A"
            Source = "Teams"
        }
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $path = "C:\Script-Export"
    if (-not (Test-Path $path)) { New-Item $path -ItemType Directory | Out-Null }
    $file = "$path\OfficeApps-$timestamp-$hostname.csv"
    $results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $file" -ForegroundColor Green
}

function Run-DriverValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportPath = "C:\Script-Export\Installed_Drivers_${hostname}_$timestamp.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Sort-Object DeviceName
    $driverInfo = $drivers | Select-Object `
        DeviceName, DeviceID, DriverVersion, DriverProviderName, InfName,
        @{Name="DriverDate";Expression={if ($_.DriverDate) {[datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null)} else { $null }}},
        Manufacturer, DriverPath
    $driverInfo | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
}

function Run-RoamingProfileValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $csvFile = "C:\Script-Export\Profiles_Applications_$hostname_$timestamp.csv"
    $profileData = @()
    $profiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }
    foreach ($p in $profiles) {
        $name = $p.LocalPath.Split('\')[-1]
        $apps = Get-ChildItem "C:\Users\$name\AppData\Local" -Directory -ErrorAction SilentlyContinue
        foreach ($a in $apps) {
            $profileData += [PSCustomObject]@{
                ProfileName = $name
                Application = $a.Name
                Path = $a.FullName
            }
        }
    }
    $profileData | Export-Csv -Path $csvFile -NoTypeInformation
    Write-Host "Exported to: $csvFile" -ForegroundColor Green
}

function Run-BrowserExtensionDetails {
    Write-Host "Standard browser extension collection not implemented in this version." -ForegroundColor Yellow
    Read-Host -Prompt "Press any key to continue"
}

function Run-OSQueryBrowserExtensions {
    Write-Host "OSQuery browser extension audit not implemented in GitHub-safe version." -ForegroundColor Yellow
    Read-Host -Prompt "Press any key to continue"
}

function Run-SSLCipherValidation {
    Write-Host "SSL Cipher validation (Nmap) not included in GitHub-safe version." -ForegroundColor Yellow
    Read-Host -Prompt "Press any key to continue"
}

function Run-WindowsPatchDetails {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME

    $out1 = "C:\Script-Export\WindowsPatches-GetHotFix-$timestamp-$hostname.csv"
    if (-not (Test-Path (Split-Path $out1))) { New-Item -ItemType Directory -Path (Split-Path $out1) | Out-Null }
    Get-HotFix | Export-Csv -Path $out1 -NoTypeInformation
    Write-Host "Patch data exported to: $out1" -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportFolder = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$exportFolder\ScriptExport_${hostname}_$timestamp.zip"

    if (-not (Test-Path $exportFolder)) {
        Write-Host "Folder '$exportFolder' not found. Exiting..." -ForegroundColor Red
        exit
    }

    $allFiles = Get-ChildItem -Path $exportFolder -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$exportFolder'. Exiting..." -ForegroundColor Yellow
        exit
    }

    $totalSizeMB = [math]::Round(($allFiles | Measure-Object Length -Sum).Sum / 1MB, 2)
    Write-Host "`nTotal size before compression: $totalSizeMB MB" -ForegroundColor Yellow

    $zipChoice = Read-Host "`nWould you like to zip the folder contents? (Y/N)"
    if ($zipChoice -notin @('Y','y')) {
        Write-Host "Zipping skipped. Exiting..." -ForegroundColor DarkGray
        exit
    }

    if (Test-Path $zipFilePath) { Remove-Item $zipFilePath -Force }
    Compress-Archive -Path "$exportFolder\*" -DestinationPath $zipFilePath

    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)
    Write-Host "`nZIP File: $zipFilePath"
    Write-Host "Size after compression: $zipSizeMB MB" -ForegroundColor Yellow

    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $company = Read-Host "Enter Company Name"
        $tenant = Read-Host "Enter Tenant Name"
        $subject = "Script Export from $hostname [$company / $tenant]"
        $body = "Attached is the export ZIP file from $hostname.`nCompany: $company`nTenant: $tenant`nZIP Path: $zipFilePath"

        $mailto = "mailto:$recipient`?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
        Start-Process $mailto
        Write-Host "`n📌 Please manually attach the ZIP file: $zipFilePath" -ForegroundColor Cyan
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Read-Host "`nPress ENTER to exit..."
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Run-ZipAndEmailResults }
        "Q" { exit }
    }
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
