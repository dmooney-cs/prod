# === Data Collection and Validation Tool ===

# Function to show main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
}

# Function to run the validation scripts menu
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
            "1" { 
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
            "2" { 
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
            "3" { 
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
            "4" { 
                Write-Host "Standard browser extension collection not implemented in this version." -ForegroundColor Yellow
                Read-Host -Prompt "Press any key to continue"
            }
            "5" { 
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $hostname = $env:COMPUTERNAME
                $outputDir = "C:\Script-Export"
                $outputCsv = "$outputDir\OSquery-browserext-$timestamp-$hostname.csv"
                $outputJson = "$outputDir\RawBrowserExt-$timestamp-$hostname.json"
                if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }
                $osqueryPaths = @("C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe", "C:\Windows\CyberCNSAgent\osqueryi.exe")
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
            "6" { 
                # Validate SSL Ciphers code here, previously provided
            }
            "7" { 
                # Windows Patch Details code here, previously provided
            }
            "8" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

# Show Main Menu and start the script loop
Show-MainMenu
$choice = Read-Host "Enter your choice"
switch ($choice) {
    "1" { Run-ValidationScripts }
    "2" { Write-Host "[Placeholder] Agent Maintenance" }
    "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
    "4" { Write-Host "[Placeholder] Zip and Email Results" }
    "Q" { exit }
}
