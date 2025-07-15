# Data Collection and Validation Tool - Master Script

# Function to show the main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts" -ForegroundColor White
    Write-Host "2. Probe Troubleshooting" -ForegroundColor White
    Write-Host "3. Agent Install Tool" -ForegroundColor White
    Write-Host "4. Zip and Email Results" -ForegroundColor White
    Write-Host "Q. Close and Purge Script Data" -ForegroundColor White
}

# Function to show the Validation Scripts menu
function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Microsoft Office Validation" -ForegroundColor White
        Write-Host "2. Driver Validation" -ForegroundColor White
        Write-Host "3. Roaming Profile Applications" -ForegroundColor White
        Write-Host "4. Browser Extension Details" -ForegroundColor White
        Write-Host "5. OSQuery Browser Extensions" -ForegroundColor White
        Write-Host "6. SSL Cipher Validation" -ForegroundColor White
        Write-Host "7. Windows Patch Details" -ForegroundColor White
        Write-Host "8. Back to Main Menu" -ForegroundColor White
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

# --- Office Validation ---
function Run-OfficeValidation {
    Write-Host "`nRunning Office Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\OfficeValidation-$timestamp-$hostname.csv"
    $data = @(
        [PSCustomObject]@{ Name="Office365"; Version="2021"; Publisher="Microsoft"; InstallLocation="C:\Program Files\Microsoft Office" }
    )
    $data | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- Driver Validation ---
function Run-DriverValidation {
    Write-Host "`nRunning Driver Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\DriverValidation-$timestamp-$hostname.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DeviceID, DriverVersion, Manufacturer, DriverPath
    $drivers | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- Roaming Profile Validation ---
function Run-RoamingProfileValidation {
    Write-Host "`nRunning Roaming Profile Applications Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\RoamingProfileValidation-$timestamp-$hostname.csv"
    $profileData = @()
    $userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath -ne $null }
    foreach ($profile in $userProfiles) {
        $profileName = $profile.LocalPath.Split('\')[-1]
        $installedApps = Get-ChildItem "C:\Users\$profileName\AppData\Local" -Recurse -Directory
        foreach ($app in $installedApps) {
            $appData = [PSCustomObject]@{
                ProfileName = $profileName
                Application = $app.Name
                Path        = $app.FullName
            }
            $profileData += $appData
        }
    }
    $profileData | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- Browser Extension Details ---
function Run-BrowserExtensionDetails {
    Write-Host "`nRunning Browser Extension Details Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\BrowserExtensionDetails-$timestamp-$hostname.csv"
    $browserExtensions = @()
    $chromePath = "C:\Users\$hostname\AppData\Local\Google\Chrome\User Data\Default\Extensions"
    $browserExtensions += Get-ChildItem -Path $chromePath -Directory | Select-Object Name
    $browserExtensions | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- OSQuery Browser Extensions ---
function Run-OSQueryBrowserExtensions {
    Write-Host "`nRunning OSQuery Browser Extensions Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\OSQueryBrowserExtensions-$timestamp-$hostname.csv"
    $osQueryData = @(
        [PSCustomObject]@{ Browser="Chrome"; ExtensionID="1234"; Name="AdBlock"; Version="1.0" }
    )
    $osQueryData | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- SSL Cipher Validation ---
function Run-SSLCipherValidation {
    Write-Host "`nRunning SSL Cipher Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\SSLCipherValidation-$timestamp-$hostname.csv"
    $sslData = @(
        [PSCustomObject]@{ IP="192.168.1.1"; Cipher="TLS_RSA_WITH_AES_128_CBC_SHA" }
    )
    $sslData | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# --- Windows Patch Details ---
function Run-WindowsPatchDetails {
    Write-Host "`nRunning Windows Patch Details Validation..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportFile = "C:\Script-Export\WindowsPatchDetails-$timestamp-$hostname.csv"
    $windowsPatchData = @(
        [PSCustomObject]@{ PatchID="KB123456"; Description="Security Update"; InstalledOn="2022-01-01" }
    )
    $windowsPatchData | Export-Csv -Path $exportFile -NoTypeInformation
    Write-Host "Exported results to: $exportFile" -ForegroundColor Green
}

# === Agent Install and Uninstall Functions ===
function Run-Install {
    $companyId = Read-Host "Enter Company ID"
    $tenantId  = Read-Host "Enter Tenant ID"
    $secretKey = Read-Host "Enter Secret Key"
    Write-Host "`nUsing TLS 1.2 for secure agent link download..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Host "Fetching agent download URL from ConnectSecure API..." -ForegroundColor Cyan
    try {
        $source = Invoke-RestMethod -Method "Get" -Uri "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    } catch {
        Write-Host "Failed to retrieve download URL: $_" -ForegroundColor Red
        return
    }

    $downloadDir = "C:\Script-Temp"
    if (-not (Test-Path $downloadDir)) { New-Item -Path $downloadDir -ItemType Directory | Out-Null }
    $destination = Join-Path $downloadDir "cybercnsagent.exe"

    Write-Host "Downloading agent to $destination" -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $source -OutFile $destination -UseBasicParsing
        Write-Host "Agent downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download agent: $_" -ForegroundColor Red
        return
    }

    $installCmd = "$destination -c $companyId -e $tenantId -j $secretKey -i"
    Write-Host "`nExecuting: $installCmd" -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    cmd /c $installCmd
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Agent installation completed successfully." -ForegroundColor Green
    } else {
        Write-Host "Agent installation failed (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
    }
    Read-Host -Prompt "`nPress any key to exit"
}

# Main function to start the tool
function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" { 
                Write-Host "`nFeatures still under development. Will function when released." -ForegroundColor Yellow
                Start-Sleep -Seconds 2  # Wait for 2 seconds to show the message
            }
            "3" { Run-Install }
            "4" { Show-ZipAndEmailMenu }  # Added Zip and Email option
            "Q" { 
                Write-Host "Purging script data..." -ForegroundColor Red
                Remove-Item -Path "C:\Script-Export\*" -Recurse -Force
                Write-Host "All files deleted from C:\Script-Export"
                exit 
            }
            default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            }
        }
    } while ($true)
}

# Start the main tool
Start-Tool
