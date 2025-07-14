# Main Menu Function
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
}

# Validation Scripts Submenu
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
            "8" { return }  # Go back to main menu
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)  # Keep looping until "Back to Main Menu" is selected
}

# Example Validation Functions (For Office Validation)
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

# Example other validation functions (Driver Validation, etc.)
function Run-DriverValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportPath = "C:\Script-Export\Installed_Drivers_${hostname}_$timestamp.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Sort-Object DeviceName
    $driverInfo = $drivers | Select-Object `
        DeviceName,
        DeviceID,
        DriverVersion,
        DriverProviderName,
        InfName,
        @{Name="DriverDate";Expression={if ($_.DriverDate) {[datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null)} else { $null }}},
        Manufacturer,
        DriverPath
    $driverInfo | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
}

# Example: Roaming Profile Validation
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

# Main menu execution loop
function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }  # Call Validation Scripts submenu
            "2" { Write-Host "[Agent Maintenance Placeholder]" }
            "3" { Write-Host "[Probe Troubleshooting Placeholder]" }
            "4" { Write-Host "[Zip and Email Placeholder]" }
            "Q" { Write-Host "Exiting..."; exit }
            default { Write-Host "Invalid option. Please choose again." -ForegroundColor Red }
        }
    } while ($choice.ToUpper() -ne "Q")  # Loop until "Q" is selected to exit
}

# Start the tool by calling the Start-Tool function
Start-Tool
