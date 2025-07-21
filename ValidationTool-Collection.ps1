# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Collection Tool               ║
# ║ Version: Beta1 | 2025-07-21                        ║
# ╚════════════════════════════════════════════════════╝

$ExportDir = "C:\Script-Export"
if (-not (Test-Path $ExportDir)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

function Pause-Script {
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    try { [void][System.Console]::ReadKey($true) } catch { Read-Host "Press Enter to continue..." }
}

function Run-OfficeValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
    $outputFile = "$exportDir\Office_Validation_$timestamp`_$hostname.csv"
    $results = @()

    $profiles = Get-ChildItem 'C:\Users' -Directory | Where-Object {
        $_.Name -notin @('All Users', 'Default', 'Default User', 'Public')
    }

    foreach ($profile in $profiles) {
        $profilePath = $profile.FullName
        $status = "Inactive"
        if (Test-Path "$profilePath\NTUSER.DAT") {
            $status = "Active"
        }

        $teamsPath = "$profilePath\AppData\Local\Microsoft\Teams\current\Teams.exe"
        if (Test-Path $teamsPath) {
            $version = (Get-Item $teamsPath).VersionInfo.FileVersion
            $results += [PSCustomObject]@{
                Profile = $profile.Name
                Status  = $status
                Product = "Microsoft Teams"
                Version = $version
                Path    = $teamsPath
            }
        }
    }

    $officeKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
    )

    foreach ($key in $officeKeys) {
        Get-ChildItem -Path $key -ErrorAction SilentlyContinue | ForEach-Object {
            $versionKey = $_.Name
            if ($versionKey -match "\\(\d+\.\d+)$") {
                $version = $Matches[1]
                $display = (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).Path
                if ($display) {
                    $results += [PSCustomObject]@{
                        Profile = "N/A"
                        Status  = "System"
                        Product = "Microsoft Office"
                        Version = $version
                        Path    = $display
                    }
                }
            }
        }
    }

    if ($results.Count -eq 0) {
        Write-Host "No Microsoft Office or Teams installations found." -ForegroundColor Yellow
    } else {
        $results | Format-Table -AutoSize
        $results | Export-Csv -Path $outputFile -NoTypeInformation
        Write-Host "`nExported to: $outputFile" -ForegroundColor Green
    }

    Read-Host "`nPress ENTER to continue"
}

function Run-DriverValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }
    $outputFile = "$outputDir\InstalledDrivers_$timestamp`_$hostname.csv"

    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer
    $drivers | Export-Csv -Path $outputFile -NoTypeInformation
    $drivers | Format-Table -AutoSize
    Write-Host "`nDrivers exported to $outputFile" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}
function Run-RoamingProfileApplications {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }
    $outputFile = "$outputDir\RoamingProfileApps_$timestamp`_$hostname.csv"
    $results = @()

    $profiles = Get-ChildItem "C:\Users" -Directory | Where-Object {
        $_.Name -notin @("All Users", "Default", "Default User", "Public")
    }

    foreach ($profile in $profiles) {
        $appPath = "$($profile.FullName)\AppData\Local\Programs"
        $type = "Old"
        if (Test-Path "$($profile.FullName)\NTUSER.DAT") { $type = "Active" }

        if (Test-Path $appPath) {
            Get-ChildItem -Path $appPath -Directory | ForEach-Object {
                $exe = Get-ChildItem $_.FullName -Filter *.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                $version = if ($exe) { (Get-Item $exe.FullName).VersionInfo.FileVersion } else { "N/A" }
                $results += [PSCustomObject]@{
                    Profile   = $profile.Name
                    Status    = $type
                    AppName   = $_.Name
                    Version   = $version
                    Path      = $_.FullName
                }
            }
        }
    }

    $results | Format-Table -AutoSize
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "`nExported to: $outputFile" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}

function Run-BrowserExtensionAudit {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
    $outputFile = "$exportDir\BrowserExtensions_$timestamp`_$hostname.csv"
    $results = @()

    $browsers = @(
        @{ Name = "Chrome"; Path = "$env:LOCALAPPDATA\Google\Chrome\User Data" },
        @{ Name = "Edge"; Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data" },
        @{ Name = "Firefox"; Path = "$env:APPDATA\Mozilla\Firefox\Profiles" }
    )

    foreach ($browser in $browsers) {
        if (Test-Path $browser.Path) {
            Get-ChildItem $browser.Path -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $manifest = Join-Path $_.FullName "manifest.json"
                if (Test-Path $manifest) {
                    $json = Get-Content $manifest -Raw | ConvertFrom-Json
                    $results += [PSCustomObject]@{
                        Browser = $browser.Name
                        Name    = $json.name
                        Version = $json.version
                        Path    = $_.FullName
                    }
                }
            }
        }
    }

    $results | Sort-Object Browser, Name | Format-Table -AutoSize
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "`nExported to: $outputFile" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}

function Run-OSQueryBrowserExtensions {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
    $csvPath = "$exportDir\OSquery-browserext-$timestamp-$hostname.csv"
    $jsonPath = "$exportDir\OSquery-browserext-$timestamp-$hostname.json"

    $osqueryPaths = @(
        "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe",
        "C:\Windows\CyberCNSAgent\osqueryi.exe"
    )

    $osquery = $osqueryPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $osquery) {
        Write-Host "❌ osqueryi.exe not found." -ForegroundColor Red
        return
    }

    $query = "SELECT name, browser_type, version, path, sha1(name || path) as unique_id FROM chrome_extensions WHERE chrome_extensions.uid IN (SELECT uid FROM users) GROUP BY unique_id;"
    $data = & $osquery --json "SELECT * FROM users;" | ConvertFrom-Json
    $output = & $osquery --json "$query" | ConvertFrom-Json

    if ($output) {
        $output | Sort-Object browser_type, name | Format-Table name, browser_type, version, path
        $output | ConvertTo-Json | Out-File $jsonPath -Encoding UTF8
        $output | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "`nExported to: $csvPath" -ForegroundColor Green
    } else {
        Write-Host "No browser extensions found via osquery." -ForegroundColor Yellow
    }
    Read-Host "`nPress ENTER to continue"
}
function Run-SSLCipherValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }
    $csvPath = "$exportDir\TLS443Scan-$hostname-$timestamp.csv"
    $logPath = "$exportDir\SSLCipherScanLog-$timestamp.csv"

    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapPath)) {
        Write-Host "❌ nmap.exe not found at expected path." -ForegroundColor Red
        return
    }

    $localIP = (Test-Connection -ComputerName (hostname) -Count 1).IPv4Address.IPAddressToString
    $cmd = "$nmapPath --script ssl-enum-ciphers -p 443 $localIP"
    Write-Host "Running: $cmd" -ForegroundColor Cyan
    $result = & $nmapPath --script ssl-enum-ciphers -p 443 $localIP

    $log = [PSCustomObject]@{
        Timestamp = (Get-Date)
        Action    = "Ran SSL cipher scan"
        Target    = $localIP
    }
    $log | Export-Csv -Path $logPath -NoTypeInformation -Append
    $result | Out-File -FilePath $csvPath
    Write-Host "`nScan complete. Results saved to $csvPath" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}

function Run-WindowsPatchDetails {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) { New-Item -Path $exportDir -ItemType Directory | Out-Null }

    $getHotfixFile = "$exportDir\GetHotFix_$timestamp`_$hostname.csv"
    $wmicFile = "$exportDir\WMIC_Hotfix_$timestamp`_$hostname.csv"
    $osqFile = "$exportDir\OSQuery_Hotfix_$timestamp`_$hostname.csv"

    Get-HotFix | Export-Csv -Path $getHotfixFile -NoTypeInformation
    Write-Host "✅ Exported Get-HotFix to $getHotfixFile"

    try {
        $wmicOutput = wmic qfe list full /format:csv
        $wmicOutput | Out-File -FilePath $wmicFile
        Write-Host "✅ Exported WMIC QFE to $wmicFile"
    } catch {
        Write-Host "⚠️ WMIC not supported or failed" -ForegroundColor Yellow
    }

    $osquery = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    if (Test-Path $osquery) {
        $query = "SELECT * FROM patches;"
        & $osquery --json "$query" | Out-File $osqFile
        Write-Host "✅ Exported OSQuery patch results to $osqFile"
    } else {
        Write-Host "❌ osquery not found for patch scan." -ForegroundColor Red
    }

    Read-Host "`nPress ENTER to continue"
}

function Run-ActiveDirectoryCollection {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "C:\Script-Export\AD_Collection_$timestamp`_$hostname.json"

    try {
        Import-Module ActiveDirectory
    } catch {
        Write-Host "❌ Active Directory module not found." -ForegroundColor Red
        return
    }

    $users = Get-ADUser -Filter * -Properties * | Select-Object Name, SamAccountName, Enabled, LastLogonDate
    $groups = Get-ADGroup -Filter * | Select-Object Name, GroupScope, Description
    $ous = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName
    $computers = Get-ADComputer -Filter * | Select-Object Name, OperatingSystem, LastLogonDate

    $collection = @{
        users     = $users
        groups    = $groups
        ous       = $ous
        computers = $computers
        count     = @{
            users     = $users.Count
            groups    = $groups.Count
            ous       = $ous.Count
            computers = $computers.Count
        }
    }

    $collection | ConvertTo-Json -Depth 5 | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "`n✅ Active Directory data exported to $outputFile" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}

function Run-AllValidations {
    Clear-Host
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "🔁 Running All Validation Scripts..." -ForegroundColor Cyan
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "`n⚠️  Some modules may produce warnings or errors — this is expected. Please allow all scripts to finish." -ForegroundColor Yellow
    Start-Sleep -Seconds 4

    Run-OfficeValidation
    Run-DriverValidation
    Run-RoamingProfileApplications
    Run-BrowserExtensionAudit
    Run-OSQueryBrowserExtensions
    Run-SSLCipherValidation
    Run-WindowsPatchDetails
    Run-ActiveDirectoryCollection

    Write-Host "`n✅ All validations complete." -ForegroundColor Green
    Read-Host "`nPress ENTER to return to the menu"
}
function Run-ZipAndEmailResults {
    $ExportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$ExportDir\ScriptExport_${hostname}_$timestamp.zip"

    $agentLogSource = "C:\Program Files (x86)\CyberCNSAgent\logs"
    $agentLogTarget = "$ExportDir\AgentLogs"
    $includeLogs = Read-Host "Include Local ConnectSecure Agent Logs @ $agentLogSource? (Y/N)"
    if ($includeLogs -in @("Y","y")) {
        if (Test-Path $agentLogSource) {
            if (-not (Test-Path $agentLogTarget)) {
                New-Item -Path $agentLogTarget -ItemType Directory | Out-Null
            }
            Copy-Item -Path "$agentLogSource\*" -Destination $agentLogTarget -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✔ Agent logs copied to: $agentLogTarget" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Agent log folder not found at: $agentLogSource" -ForegroundColor Yellow
        }
    }

    if (-not (Test-Path $ExportDir)) {
        Write-Host "❌ Folder '$ExportDir' not found." -ForegroundColor Red
        Pause-Script
        return
    }

    $allFiles = Get-ChildItem -Path $ExportDir -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "⚠️ No files found in '$ExportDir'." -ForegroundColor Yellow
        Pause-Script
        return
    }

    Compress-Archive -Path "$ExportDir\*" -DestinationPath $zipFilePath -Force
    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)
    Write-Host "`n✅ ZIP created: $zipFilePath ($zipSizeMB MB)" -ForegroundColor Green

    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $bodyFallback = "Please attach the file located at: $zipFilePath before sending this email."

        $modernOutlookPath = "$env:LOCALAPPDATA\Packages\microsoft.windowscommunicationsapps_8wekyb3d8bbwe"
        $isModernOutlook = Test-Path $modernOutlookPath

        Write-Host "`nChecking for Outlook..." -ForegroundColor Cyan
        if ($isModernOutlook) {
            Write-Host "🧭 New Outlook (Microsoft Store version) detected. COM automation is not supported." -ForegroundColor Yellow
        }

        try {
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            try { $Outlook = New-Object -ComObject Outlook.Application } catch { $Outlook = $null }
        }

        if ($Outlook) {
            try {
                $namespace = $Outlook.GetNamespace("MAPI")
                $namespace.Logon($null, $null, $false, $true)
                $Mail = $Outlook.CreateItem(0)
                $Mail.To = $recipient
                $Mail.Subject = $subject
                $Mail.Body = "Attached is the export ZIP file from $hostname.`nZIP Path: $zipFilePath"
                if (Test-Path $zipFilePath) {
                    $Mail.Attachments.Add($zipFilePath)
                }
                $Mail.Display()
                Write-Host "`n✅ Outlook draft email opened with ZIP attached." -ForegroundColor Green
            } catch {
                Write-Host "❌ Outlook COM error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "`n⚠️ Outlook COM interface not available. Launching default mail client..." -ForegroundColor Yellow
            if ($isModernOutlook) {
                Write-Host "✳️ This appears to be the New Outlook (Microsoft Store version), which cannot auto-attach files." -ForegroundColor Magenta
            }

            $mailto = "mailto:$recipient?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($bodyFallback))"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"$mailto`"" -WindowStyle Hidden
            Write-Host "`nPlease manually attach this file:" -ForegroundColor Cyan
            Write-Host "$zipFilePath" -ForegroundColor White
        }
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Pause-Script
}

function Run-CleanupScriptData {
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) {
        Write-Host "`n⚠️ Export folder not found at $exportDir" -ForegroundColor Yellow
        Pause-Script
        return
    }

    $files = Get-ChildItem -Path $exportDir -Recurse -Force -File
    if ($files.Count -eq 0) {
        Write-Host "`nℹ️ No export files to delete." -ForegroundColor Cyan
        Pause-Script
        return
    }

    Write-Host "`n⚠️ The following files will be deleted from $exportDir:" -ForegroundColor Red
    $files | ForEach-Object { Write-Host $_.FullName -ForegroundColor DarkGray }

    $confirm = Read-Host "`nType 'DELETE' to confirm"
    if ($confirm -eq "DELETE") {
        $files | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Host "`n✅ Export data cleared." -ForegroundColor Green
    } else {
        Write-Host "`nCleanup cancelled." -ForegroundColor Yellow
    }

    Pause-Script
}

function Show-CollectionMenu {
    do {
        Write-Host "`n========= 🧩 Collection Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Office Validation"
        Write-Host "[2] Driver Validation"
        Write-Host "[3] Roaming Profile Applications"
        Write-Host "[4] Browser Extension Details"
        Write-Host "[5] OSQuery Browser Extensions"
        Write-Host "[6] SSL Cipher Validation"
        Write-Host "[7] Windows Patch Details"
        Write-Host "[8] Active Directory Collection"
        Write-Host "[9] Run All Validation Scripts"
        Write-Host "[10] Zip and Email Results"
        Write-Host "[11] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileApplications }
            "4" { Run-BrowserExtensionAudit }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { Run-ActiveDirectoryCollection }
            "9" { Run-AllValidations }
            "10" { Run-ZipAndEmailResults }
            "11" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-CollectionMenu
