
function Run-WindowsPatchDetails {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME

    Write-Host "`n[1/3] Collecting patch info via Get-HotFix..." -ForegroundColor Cyan
    $hotfixFolder = "C:\Script-Export\HotFix-Report"
    if (-not (Test-Path $hotfixFolder)) {
        New-Item -ItemType Directory -Path $hotfixFolder | Out-Null
    }
    $getHotfixFile = "$hotfixFolder\GetHotFix-$timestamp-$hostname.csv"
    try {
        $getHotfixes = Get-HotFix | Where-Object { $_.HotFixID } | ForEach-Object {
            [PSCustomObject]@{
                HotfixID = $_.HotFixID
                Description = $_.Description
                InstalledOn = $_.InstalledOn
                InstalledBy = $_.InstalledBy
                Source = "Get-HotFix"
            }
        }
        if ($getHotfixes.Count -gt 0) {
            $getHotfixes | Export-Csv -Path $getHotfixFile -NoTypeInformation -Encoding UTF8
            Write-Host "Get-HotFix results saved to:" -ForegroundColor Green
            Write-Host "$getHotfixFile" -ForegroundColor Cyan
        } else {
            Write-Host "No results from Get-HotFix." -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Failed to retrieve results via Get-HotFix: $_"
    }

    Write-Host "`n[2/3] Collecting patch info via WMIC..." -ForegroundColor Cyan
    $wmicOutputFolder = "C:\Script-Export\WMIC-Patch-Report"
    if (-not (Test-Path $wmicOutputFolder)) {
        New-Item -ItemType Directory -Path $wmicOutputFolder | Out-Null
    }
    $wmicOutputFile = "$wmicOutputFolder\WMIC-Patches-$timestamp-$hostname.csv"
    $wmicHotfixes = @()
    try {
        $wmicRaw = (wmic qfe get HotfixID | findstr /v HotFixID) -split "`r`n"
        $wmicHotfixes = $wmicRaw | Where-Object { $_ -and $_.Trim() -ne "" } | ForEach-Object {
            [PSCustomObject]@{
                HotfixID = $_.Trim()
                Source   = "WMIC"
            }
        }
        if ($wmicHotfixes.Count -gt 0) {
            $wmicHotfixes | Export-Csv -Path $wmicOutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "WMIC patch list saved to:" -ForegroundColor Green
            Write-Host "$wmicOutputFile" -ForegroundColor Cyan
        } else {
            Write-Host "No hotfixes found using WMIC." -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "WMIC not available or failed: $_"
    }

    Write-Host "`n[3/3] Collecting patch info via osquery..." -ForegroundColor Cyan
    $osqueryPath = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    $osqueryOutputFolder = "C:\Script-Export\OSQuery-Patch-Report"
    if (-not (Test-Path $osqueryOutputFolder)) {
        New-Item -ItemType Directory -Path $osqueryOutputFolder | Out-Null
    }
    $osqueryOutputFile = "$osqueryOutputFolder\OSQuery-Patches-$timestamp-$hostname.csv"

    if (-not (Test-Path $osqueryPath)) {
        Write-Host "ERROR: osqueryi.exe not found at: $osqueryPath" -ForegroundColor Red
    } else {
        $queries = @(
            "select  CONCAT('KB',replace(split(split(title, 'KB',1),' ',0),')','')) as hotfix_id,description, datetime(date,'unixepoch') as install_date,'' as installed_by,'' as installed_on from windows_update_history where title like '%KB%' group by split(split(title, 'KB',1),' ',0);",
            "select hotfix_id,description,installed_by,install_date,installed_on from patches group by hotfix_id;"
        )
        $results = @()
        foreach ($query in $queries) {
            $output = & "$osqueryPath" --json "$query" 2>$null
            if ($output) {
                try {
                    $parsed = $output | ConvertFrom-Json
                    $results += $parsed
                } catch {
                    Write-Warning "Failed to parse osquery JSON output: $_"
                }
            }
        }
        if ($results.Count -gt 0) {
            $results | Export-Csv -Path $osqueryOutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "OSQuery patch list saved to:" -ForegroundColor Green
            Write-Host "$osqueryOutputFile" -ForegroundColor Cyan
        } else {
            Write-Host "No osquery patch data found or parsed." -ForegroundColor Yellow
        }
    }

    Write-Host "`nPatch collection complete." -ForegroundColor Magenta
    Read-Host -Prompt "Press Enter to return"
}


function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Browser Extension details"
        Write-Host "4. SSL Cipher Validation"
        Write-Host "5. Windows Patch Details"
        Write-Host "6. Back to Main Menu"
        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-BrowserExtensionDetails }
            "4" { Run-SSLCipherValidation }
            "5" { Run-WindowsPatchDetails }
            "6" { return }
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



