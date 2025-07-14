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

# Microsoft Office Validation Script
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
    Read-Host -Prompt "Press any key to continue"
}

# Driver Validation Script
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
        @{Name="DriverDate";Expression={
            if ($_.DriverDate) { 
                [datetime]::ParseExact($_.DriverDate, 'yyyyMMddHHmmss.000000+000', $null) 
            } else { 
                $null 
            }}},
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
            }}},
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

    $driverInfo | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nInstalled driver summary exported to:" -ForegroundColor Cyan
    Write-Host $exportPath -ForegroundColor Green
    Read-Host -Prompt "Press any key to continue"
}

# Browser Extension Details Script
function Run-BrowserExtensionDetails {
    $AllResults = @()
    $Users = Get-ChildItem 'C:\Users' -Directory | Where-Object { $_.Name -notin @('Default', 'Public', 'All Users') }

    foreach ($user in $Users) {
        $profilePath = $user.FullName
        $chromePath = Join-Path $profilePath 'AppData\Local\Google\Chrome\User Data\Default\Extensions'
        $AllResults += Get-ChromeEdgeExtensions -BrowserName 'Chrome' -BasePath $chromePath

        $edgePath = Join-Path $profilePath 'AppData\Local\Microsoft\Edge\User Data\Default\Extensions'
        $AllResults += Get-ChromeEdgeExtensions -BrowserName 'Edge' -BasePath $edgePath

        $AllResults += Get-FirefoxExtensions -UserProfile $profilePath
    }

    $SortedResults = $AllResults | Sort-Object Browser

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportPath = "C:\Script-Export"
    if (-not (Test-Path $exportPath)) { New-Item $exportPath -ItemType Directory }

    $csvPath = Join-Path $exportPath "BrowserExtensions_$timestamp`_$hostname.csv"
    $SortedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport saved to: $csvPath" -ForegroundColor Green
    Read-Host -Prompt "Press any key to continue"
}

# SSL Cipher Validation Script
function Run-SSLCipherValidation {
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $logFile = "$exportDir\SSLCipherScanLog-$timestamp.csv"
    $outputFile = "$exportDir\TLS443Scan-$hostname-$timestamp.csv"

    function Log-Action {
        param([string]$Action)
        [PSCustomObject]@{
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Action    = $Action
        } | Export-Csv -Path $logFile -NoTypeInformation -Append
    }

    Log-Action "Starting SSL Cipher Scan"

    $localIP = (Get-NetIPAddress -AddressFamily IPv4 |
        Where-Object { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' } |
        Select-Object -First 1 -ExpandProperty IPAddress)

    if (-not $localIP) {
        Write-Host "❌ Could not determine local IP address." -ForegroundColor Red
        Log-Action "Error: Could not determine local IP address"
        Start-Sleep -Seconds 2
        return
    }

    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap"
    $nmapExe = Join-Path $nmapPath "nmap.exe"
    $npcapInstaller = Join-Path $nmapPath "npcap-1.50-oem.exe"

    if (-not (Test-Path $nmapExe)) {
        Write-Host "`n❌ Nmap not found. Please convert the ConnectSecure agent to a probe, and re-attempt this scan." -ForegroundColor Red
        Write-Host "ℹ️  In the future, NMAP will be automatically installed if missing." -ForegroundColor Yellow
        Log-Action "Error: Nmap not found at expected path"
        Start-Sleep -Seconds 2
        return
    }

    function Test-Npcap {
        $versionOutput = & $nmapExe -V 2>&1
        return ($versionOutput -match "Could not import all necessary Npcap functions")
    }

    Write-Host "`n🔍 Testing Nmap for Npcap compatibility..." -ForegroundColor Cyan
    Log-Action "Checking for Npcap compatibility via Nmap"

    if (Test-Npcap) {
        Write-Host "`n⚠️  Npcap not functioning or incomplete." -ForegroundColor Yellow
        Log-Action "Npcap issue detected"

        if (Test-Path $npcapInstaller) {
            Write-Host "⚙️  Installing bundled Npcap silently with WinPcap compatibility..." -ForegroundColor Cyan
            Log-Action "Installing Npcap from $npcapInstaller"

            Start-Process -FilePath $npcapInstaller -ArgumentList "/S", "/winpcap_mode=yes" -Wait -NoNewWindow

            Write-Host "`n✅ Npcap installed. Retesting Nmap..." -ForegroundColor Green
            Log-Action "Npcap installed. Retesting Nmap compatibility."

            if (Test-Npcap) {
                Write-Host "❌ Npcap issue still detected after installation. Scan may be incomplete." -ForegroundColor Red
                Log-Action "Warning: Npcap issue persists after install"
            }
        } else {
            Write-Host "`n❌ Required installer not found: $npcapInstaller" -ForegroundColor Red
            Write-Host "➡️  Convert agent to Probe before attempting." -ForegroundColor Yellow
            Log-Action "Npcap installer missing. User prompted to convert agent to probe."
            Start-Sleep -Seconds 2
            return
        }
    }

    Write-Host "`n🔍 Running SSL Cipher enumeration on $localIP:443..." -ForegroundColor Cyan
    Log-Action "Running Nmap SSL scan on $localIP"

    try {
        $scanResult = & $nmapExe --script ssl-enum-ciphers -p 443 $localIP 2>&1
        $scanResult | Set-Content -Path $outputFile -Encoding UTF8

        if ($scanResult -match "Could not import all necessary Npcap functions") {
            Write-Host "`n⚠️  Nmap still reports Npcap issue after install. Results may be incomplete." -ForegroundColor Yellow
            Log-Action "Npcap warning detected in scan output"
        } else {
            Write-Host "`n=== SSL Cipher Scan Results ===" -ForegroundColor White
            $scanResult
            Log-Action "SSL Cipher scan completed successfully"
        }
    } catch {
        Write-Host "❌ Nmap scan failed: $_" -ForegroundColor Red
        Log-Action "Error: Nmap scan failed"
        Start-Sleep -Seconds 2
        return
    }

    Write-Host "`n✅ Scan complete. Results saved to:" -ForegroundColor Green
    Write-Host "$outputFile" -ForegroundColor Green
    Log-Action "Results saved to $outputFile"
    Log-Action "SSL Cipher Scan completed"

    Read-Host -Prompt "Press ENTER to exit..."
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" { Write-Host "[Placeholder] Agent Maintenance" }
            "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
            "4" { Run-ZipAndEmailResults }
            "Q" { exit }
        }
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
