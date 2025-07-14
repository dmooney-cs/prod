
# Data-Collection-and-Validation-Tool.ps1
# Complete Self-contained PowerShell Script for GitHub Deployment

Write-Host "`n======== Data Collection and Validation Tool ========" -ForegroundColor Cyan

function Show-MainMenu {
    Write-Host ""
    Write-Host "1. Validation Scripts"
    Write-Host "   → Run application, driver, browser, SSL, and update validations."
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
        "2" { Write-Host "[Simulated] Agent Maintenance..." -ForegroundColor Cyan; Pause }
        "3" { Write-Host "[Simulated] Probe Troubleshooting..." -ForegroundColor Cyan; Pause }
        "4" { Write-Host "[Simulated] Zipping and emailing results..." -ForegroundColor Cyan; Pause }
        "Q" { Purge-ScriptData }
        default { Write-Host "Invalid option. Please select again." -ForegroundColor Red }
    }
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Browser Extension details"
        Write-Host "4. SSL Cipher Validation"
        Write-Host "5. Back to Main Menu"
        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-BrowserExtensionDetails }
            "4" { Run-SSLCipherValidation }
            "5" { return }
            default { Write-Host "Invalid choice." -ForegroundColor Red }
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
            if ($_.DisplayName -and ($appFilter -eq "" -or $_.DisplayName -like "*$appFilter*")) {
                $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }
                $results += [PSCustomObject]@{
                    Name = $_.DisplayName; Version = $_.DisplayVersion; Publisher = $_.Publisher
                    InstallLocation = $installLocation; InstallDate = $_.InstallDate
                    InstalledBy = $_.InstallSource; Source = "Registry"; Category = "Application"
                    RegistryPath = $_.PSPath -replace "^Microsoft.PowerShell.Core\\Registry::", ""
                    ProfileType = "N/A"
                }
            }
        }
    }

    Get-AppxPackage | ForEach-Object {
        if ($appFilter -eq "" -or $_.Name -like "*$appFilter*" -or $_.PackageFullName -like "*$appFilter*") {
            $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }
            $results += [PSCustomObject]@{
                Name = $_.Name; Version = $_.Version; Publisher = $_.Publisher
                InstallLocation = $installLocation; InstallDate = "N/A"; InstalledBy = "N/A"
                Source = "Microsoft Store"; Category = "Microsoft Store App"
                RegistryPath = "N/A"; ProfileType = "N/A"
            }
        }
    }

    $teamsPaths = Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Teams\Teams.exe" -Recurse -ErrorAction SilentlyContinue
    foreach ($teamsPath in $teamsPaths) {
        $results += [PSCustomObject]@{
            Name = 'Microsoft Teams'; Version = (Get-Item $teamsPath.FullName).VersionInfo.FileVersion
            Publisher = 'Microsoft'; InstallLocation = $teamsPath.DirectoryName
            InstallDate = "N/A"; InstalledBy = "N/A"; Source = "Teams (Local)"
            Category = "Microsoft Store App"; RegistryPath = "N/A"; ProfileType = "N/A"
        }
    }

    if ($results.Count -eq 0) {
        Write-Host "No matching applications found." -ForegroundColor Yellow
    } else {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $hostname = $env:COMPUTERNAME
        $exportPath = "C:\Script-Export"
        if (-not (Test-Path -Path $exportPath)) {
            New-Item -Path $exportPath -ItemType Directory | Out-Null
        }
        $csvPath = "$exportPath\OfficeApps_$timestamp`_$hostname.csv"
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
        DeviceName, DeviceID, DriverVersion, DriverProviderName, InfName,
        @{Name="DriverDate";Expression={if ($_.DriverDate) { [datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null) } else { $null }}},
        Manufacturer, DriverPath,
        @{Name="BestGuessDriverFullPath";Expression={
            $file = [System.IO.Path]::GetFileName($_.DriverPath)
            $paths = @("C:\Windows\System32\drivers\$file", "C:\Windows\System32\DriverStore\FileRepository\$file")
            foreach ($p in $paths) { if (Test-Path $p) { return $p } }; return "Unknown"
        }},
        @{Name="FullInfFilePath";Expression={
            $inf = $_.InfName
            $paths = @("C:\Windows\INF\$inf", "C:\Windows\System32\DriverStore\FileRepository\$inf")
            foreach ($p in $paths) { if (Test-Path $p) { return $p } }; return "INF not found"
        }}
    $driverInfo | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nInstalled driver summary exported to: $outputPath" -ForegroundColor Green
    Read-Host -Prompt "Press Enter to continue"
}

function Run-BrowserExtensionDetails {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    $outputCsv = "$outputDir\OSquery-browserext-$timestamp-$hostname.csv"
    $outputJson = "$outputDir\RawBrowserExt-$timestamp-$hostname.json"
    if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }
    $osqueryPaths = @("C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe", "C:\Windows\CyberCNSAgent\osqueryi.exe")
    $osquery = $osqueryPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $osquery) {
        Write-Host "❌ osqueryi.exe not found." -ForegroundColor Red; return
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
    $parsed = $json | ConvertFrom-Json
    if ($parsed.Count -eq 0) {
        Write-Host "⚠️ No browser extensions found." -ForegroundColor Yellow
    } else {
        $parsed | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
        Write-Host "`nBrowser extensions exported to: $outputCsv" -ForegroundColor Green
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Run-SSLCipherValidation {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
    $outputFile = "$exportDir\TLS443Scan-$hostname-$timestamp.csv"
    $nmapExe = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapExe)) {
        Write-Host "❌ Nmap not found. Please convert agent to probe." -ForegroundColor Red; return
    }
    $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' } | Select-Object -First 1 -ExpandProperty IPAddress)
    if (-not $localIP) { Write-Host "❌ Could not determine local IP." -ForegroundColor Red; return }
    $scanResult = & $nmapExe --script ssl-enum-ciphers -p 443 $localIP 2>&1
    $scanResult | Set-Content -Path $outputFile -Encoding UTF8
    Write-Host "`nSSL Cipher scan complete. Output saved to: $outputFile" -ForegroundColor Green
    Read-Host -Prompt "Press ENTER to continue"
}

function Purge-ScriptData {
    $tempPath = "C:\Script-Temp"
    if (Test-Path $tempPath) {
        Remove-Item -Path $tempPath\* -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Temp folder purged." -ForegroundColor Cyan
    }
    Write-Host "Press Enter to exit..."; Read-Host | Out-Null; exit
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
