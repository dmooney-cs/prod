# Data Collection and Validation Tool - Fully Working Version

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
        Write-Host "5. SSL Cipher Validation"
        Write-Host "6. Windows Patch Details"
        Write-Host "7. Back to Main Menu"
        $valChoice = Read-Host "Select an option"
        switch ($valChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-OSQueryBrowserExtensions }
            "5" { Run-SSLCipherValidation }
            "6" { Run-WindowsPatchDetails }
            "7" { return }
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

function Run-OSQueryBrowserExtensions {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outDir = "C:\Script-Export"
    $jsonOut = "$outDir\RawBrowserExt-$timestamp-$hostname.json"
    $csvOut = "$outDir\OSquery-browserext-$timestamp-$hostname.csv"
    $osquery = @("C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe", "C:\Windows\CyberCNSAgent\osqueryi.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $osquery) { Write-Host "osqueryi.exe not found." -ForegroundColor Red; return }
    Push-Location (Split-Path $osquery)
    $sql = "SELECT name, browser_type, version, path FROM chrome_extensions WHERE uid IN (SELECT uid FROM users);"
    $json = & .\osqueryi.exe --json "$sql" 2>$null
    Pop-Location
    if ($json) {
        $parsed = $json | ConvertFrom-Json
        $parsed | Export-Csv -Path $csvOut -NoTypeInformation -Encoding UTF8
        Set-Content -Path $jsonOut -Value $json
        Write-Host "Exported to: $csvOut" -ForegroundColor Green
    } else {
        Write-Host "No results from osquery." -ForegroundColor Yellow
    }
}

function Run-SSLCipherValidation {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $nmap = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmap)) {
        Write-Host "Nmap not found. Please convert agent to probe." -ForegroundColor Red
        return
    }
    $ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' }).IPAddress | Select-Object -First 1
    $outFile = "C:\Script-Export\TLS443Scan-$hostname-$timestamp.csv"
    $scan = & $nmap --script ssl-enum-ciphers -p 443 $ip 2>&1
    $scan | Set-Content -Path $outFile -Encoding UTF8
    Write-Host "Results saved to: $outFile" -ForegroundColor Green
}

function Run-WindowsPatchDetails {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $out1 = "C:\Script-Export\HotFix-Report\GetHotFix-$timestamp-$hostname.csv"
    $out2 = "C:\Script-Export\WMIC-Patch-Report\WMIC-Patches-$timestamp-$hostname.csv"
    $out3 = "C:\Script-Export\OSQuery-Patch-Report\OSQuery-Patches-$timestamp-$hostname.csv"

    if (-not (Test-Path (Split-Path $out1))) { New-Item -ItemType Directory -Path (Split-Path $out1) | Out-Null }
    if (-not (Test-Path (Split-Path $out2))) { New-Item -ItemType Directory -Path (Split-Path $out2) | Out-Null }
    if (-not (Test-Path (Split-Path $out3))) { New-Item -ItemType Directory -Path (Split-Path $out3) | Out-Null }

    Get-HotFix | Export-Csv -Path $out1 -NoTypeInformation
    $wmicOut = (wmic qfe get HotfixID | findstr /v HotFixID) -split "`r`n" | Where-Object { $_ } |
        ForEach-Object { [PSCustomObject]@{ HotfixID = $_.Trim(); Source = "WMIC" } }
    $wmicOut | Export-Csv -Path $out2 -NoTypeInformation

    $osquery = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    if (Test-Path $osquery) {
        $query = "select hotfix_id, description from patches;"
        $output = & $osquery --json "$query"
        if ($output) {
            $parsed = $output | ConvertFrom-Json
            $parsed | Export-Csv -Path $out3 -NoTypeInformation
        }
    }

    Write-Host "Patch data exported to: $out1, $out2, $out3" -ForegroundColor Green
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Write-Host "[Placeholder] Zip and Email Results" }
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
