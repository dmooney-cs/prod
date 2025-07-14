# Data Collection and Validation Tool - Master Script

# Function to show the main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
}

# Function to show the Validation Scripts menu
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

# Office Validation
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

# Driver Validation
function Run-DriverValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportPath = "C:\Script-Export\Installed_Drivers_${hostname}_$timestamp.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Sort-Object DeviceName
    $driverInfo = $drivers | Select-Object `DeviceName, DeviceID, DriverVersion, DriverProviderName, InfName, 
        @{Name="DriverDate";Expression={if ($_.DriverDate) {[datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null)} else { $null }}},
        Manufacturer, DriverPath
    $driverInfo | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
}

# Zip and Email Results
function Run-ZipAndEmailResults {
    # Set up folder and output paths
    $exportFolder = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$exportFolder\ScriptExport_${hostname}_$timestamp.zip"

    # Ensure export folder exists
    if (-not (Test-Path $exportFolder)) {
        Write-Host "Folder '$exportFolder' not found. Exiting..." -ForegroundColor Red
        exit
    }

    # Get all files recursively
    $allFiles = Get-ChildItem -Path $exportFolder -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$exportFolder'. Exiting..." -ForegroundColor Yellow
        exit
    }

    # Display contents
    Write-Host ""
    Write-Host "=== Contents of $exportFolder ===" -ForegroundColor Cyan
    $allFiles | ForEach-Object { Write-Host $_.FullName }

    # Calculate total size before compression
    $totalSizeMB = [math]::Round(($allFiles | Measure-Object Length -Sum).Sum / 1MB, 2)
    Write-Host ""
    Write-Host "Total size before compression: $totalSizeMB MB" -ForegroundColor Yellow

    # Prompt to ZIP
    $zipChoice = Read-Host "`nWould you like to zip the folder contents? (Y/N)"
    if ($zipChoice -notin @('Y','y')) {
        Write-Host "Zipping skipped. Exiting..." -ForegroundColor DarkGray
        exit
    }

    # Remove old zip if exists
    if (Test-Path $zipFilePath) { Remove-Item $zipFilePath -Force }

    # Create ZIP
    Compress-Archive -Path "$exportFolder\*" -DestinationPath $zipFilePath

    # Get ZIP size
    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)

    # Show summary
    Write-Host ""
    Write-Host "=== ZIP Summary ===" -ForegroundColor Green
    Write-Host "ZIP File: $zipFilePath"
    Write-Host "Size after compression: $zipSizeMB MB" -ForegroundColor Yellow

    # Prompt to email
    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $body = "Attached is the export ZIP file from $hostname.`nZIP Path: $zipFilePath"

        Write-Host "`nChecking for Outlook..." -ForegroundColor Cyan
        $Outlook = $null
        $OutlookWasRunning = $false

        # Attempt to connect to Outlook (running or start new)
        try {
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $OutlookWasRunning = $true
        } catch {
            try {
                $Outlook = New-Object -ComObject Outlook.Application
                $OutlookWasRunning = $false
            } catch {
                $Outlook = $null
            }
        }

        if ($Outlook) {
            try {
                # Ensure MAPI session is open
                $namespace = $Outlook.GetNamespace("MAPI")
                $namespace.Logon($null, $null, $false, $true)

                # Create and populate mail item
                $Mail = $Outlook.CreateItem(0)
                $Mail.To = $recipient
                $Mail.Subject = $subject
                $Mail.Body = $body

                if (Test-Path $zipFilePath) {
                    $Mail.Attachments.Add($zipFilePath)
                } else {
                    Write-Host "❌ ZIP file not found at $zipFilePath" -ForegroundColor Red
                }

                $Mail.Display()
                Write-Host "`n✅ Outlook draft email opened with ZIP attached." -ForegroundColor Green
            } catch {
                Write-Host "❌ Outlook COM error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "`n⚠️ Outlook not available. Falling back to default mail client..." -ForegroundColor Yellow
            $mailto = "mailto:$recipient`?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
            Start-Process $mailto
            Write-Host ""
            Write-Host "Please manually attach this file:" -ForegroundColor Cyan
            Write-Host "$zipFilePath" -ForegroundColor White
        }
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Write-Host ""
    Read-Host "Press ENTER to exit..."
}

# Main function to start the tool
function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" { Write-Host "[Placeholder] Agent Maintenance" }
            "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
            "4" { Run-ZipAndEmailResults }
            "Q" { 
                # Purge script data and exit
                Write-Host "Purging script data..." -ForegroundColor Red
                Remove-Item -Path "C:\Script-Export\*" -Recurse -Force
                Write-Host "All files deleted from C:\Script-Export"
                exit 
            }
        }
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
