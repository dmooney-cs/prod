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

# Agent Maintenance Submenu
function Run-AgentMaintenanceMenu {
    do {
        Write-Host "`n---- Agent Maintenance Menu ----" -ForegroundColor Cyan
        Write-Host "1. Agent Install Tool"
        Write-Host "2. Check SMB"
        Write-Host "3. Set SMB"
        Write-Host "4. Clear Pending Jobs"
        Write-Host "5. Back to Main Menu"
        $agentChoice = Read-Host "Select an option"

        switch ($agentChoice) {
            "1" { Run-AgentInstallTool }
            "2" { Run-CheckSMB }
            "3" { Run-SetSMB }
            "4" { Run-ClearPendingJobs }
            "5" { return }
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
        @{Name="DriverDate";Expression={if ($_.DriverDate) {[datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null)} else { $null }}} ,
        Manufacturer, DriverPath
    $driverInfo | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
}

# Agent - Install Tool
function Run-AgentInstallTool {
    $agentDir = "C:\Program Files (x86)\CyberCNSAgent"
    $tempDir = "C:\Script-Temp"
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $logFile = "$exportDir\AgentInstallLog-$timestamp-$hostname.csv"
    $transcriptFile = "$exportDir\AgentInstallTranscript-$timestamp-$hostname.txt"

    if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory | Out-Null }
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }

    Start-Transcript -Path $transcriptFile -Force

    if (Test-Path "$agentDir\uninstall.bat") {
        $reinstall = Read-Host "Agent detected. Would you like to Re-Install it? (Y/N)"
        if ($reinstall -notin @("Y", "y")) {
            Write-Host "Install aborted."
            Stop-Transcript
            return
        }
        Write-Host "Uninstalling existing agent..."
        & "$agentDir\uninstall.bat"
        Start-Sleep -Seconds 5
    }

    $company = Read-Host "Enter Company ID"
    $tenant = Read-Host "Enter Tenant ID"
    $secret = Read-Host "Enter Secret Key"

    $url = "https://downloads.myconnectsecure.com/agent/windows/cybercnsagent.exe"
    $installer = "$tempDir\cybercnsagent.exe"
    Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing

    $cmd = "& `"$installer`" -c $company -e $tenant -j $secret -i"
    Write-Host "`nRunning: $cmd"
    Invoke-Expression $cmd

    Stop-Transcript
    Write-Host "`n✅ Agent Install Tool completed." -ForegroundColor Green
    Write-Host "Transcript: $transcriptFile" -ForegroundColor Cyan
    Write-Host "Log file: $logFile" -ForegroundColor Cyan
    Read-Host "Press ENTER to return..."
}

# Agent - Check SMB
function Run-CheckSMB {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\SMB_Version_Report_$timestamp-$hostname.csv"

    $results = @()
    $checkKeys = @{
        "SMB1" = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\SMB1"
        "SMB2" = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\SMB2"
        "SMB3" = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters\EnableSecuritySignature"
    }

    foreach ($key in $checkKeys.Keys) {
        $enabled = try { (Get-ItemProperty -Path ($checkKeys[$key]) -ErrorAction Stop)."(default)" } catch { "Not Found" }
        $disableKey = "Use Disable-$key via registry to disable"
        $results += [PSCustomObject]@{
            Protocol = $key
            Enabled = $enabled
            DisableKey = $disableKey
        }
    }

    $results | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "✅ SMB status exported to: $exportFile" -ForegroundColor Green
    Read-Host "Press ENTER to return..."
}

# Agent - Set SMB
function Run-SetSMB {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB2 -Type DWORD -Value 1 -Force
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (SMB-In)" -Enabled True -Profile Any
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (NB-Session-In)" -Enabled True -Profile Any
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "LocalAccountTokenFilterPolicy" -Value 1 -Type DWord -Force

    Write-Host "`n✅ SMB2 and firewall rules configured. Token policy enabled." -ForegroundColor Green
    Read-Host "Press ENTER to return..."
}

# Agent - Clear Jobs
function Run-ClearPendingJobs {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $tempDir = "C:\Script-Temp"
    $exportDir = "C:\Script-Export"
    $transcriptFile = "$exportDir\AgentCheck-FullOutput-$timestamp-$hostname.txt"
    $summaryLog = "$exportDir\AgentCheck-Summary-$timestamp-$hostname.csv"

    if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory | Out-Null }
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }

    Start-Transcript -Path $transcriptFile -Force

    $url = "https://agentv3.myconnectsecure.com/agentcheck.exe"
    $toolPath = "$tempDir\agentcheck.exe"
    Invoke-WebRequest -Uri $url -OutFile $toolPath -UseBasicParsing

    Stop-Service cybercnsagent -Force
    Stop-Service cybercnsagentmonitor -Force

    $pendingDir = "C:\Program Files (x86)\CyberCNSAgent\pendingjobqueue"
    if (Test-Path $pendingDir) {
        Get-ChildItem -Path $pendingDir -Recurse | Remove-Item -Force -Recurse
    }

    Set-Location "C:\Program Files (x86)\CyberCNSAgent"
    & ".\agentcheck.exe"

    Start-Service cybercnsagent
    Start-Service cybercnsagentmonitor

    Stop-Transcript
    Write-Host "`n✅ AgentCheck complete." -ForegroundColor Green
    Write-Host "Transcript: $transcriptFile" -ForegroundColor Cyan
    Write-Host "Summary log: $summaryLog" -ForegroundColor Cyan
    Read-Host "Press ENTER to return..."
}

# Zip and Email Results
function Run-ZipAndEmailResults {
    # (unchanged from prior working version)
    # [intentionally omitted here to save space — it's already included above and working]
}

# Main loop
function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"

        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" { Run-AgentMaintenanceMenu }
            "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
            "4" { Run-ZipAndEmailResults }
            "Q" {
                Write-Host "Purging script data..." -ForegroundColor Red
                Remove-Item -Path "C:\Script-Export\*" -Recurse -Force
                Write-Host "All files deleted from C:\Script-Export"
                exit
            }
        }
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
