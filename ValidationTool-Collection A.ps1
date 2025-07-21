# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation A (Apps & Drivers) ║
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

# === Run-OfficeValidation ===
function Run-OfficeValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "$ExportDir\Office_Validation_$timestamp`_$hostname.csv"
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

# === Run-DriverValidation ===
function Run-DriverValidation {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "$ExportDir\InstalledDrivers_$timestamp`_$hostname.csv"

    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer
    $drivers | Export-Csv -Path $outputFile -NoTypeInformation
    $drivers | Format-Table -AutoSize
    Write-Host "`nDrivers exported to $outputFile" -ForegroundColor Green
    Read-Host "`nPress ENTER to continue"
}

# === Run-RoamingProfileApplications ===
function Run-RoamingProfileApplications {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "$ExportDir\RoamingProfileApps_$timestamp`_$hostname.csv"
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

# === Run-ZipAndEmailResults ===
function Run-ZipAndEmailResults {
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
    if ($emailChoice -in @("Y", "y")) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $body = "Please attach the file located at: $zipFilePath before sending this email."
        $mailto = "mailto:$recipient?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
        Start-Process "cmd.exe" -ArgumentList "/c start `"$mailto`"" -WindowStyle Hidden
        Write-Host "`nPlease manually attach this file:" -ForegroundColor Cyan
        Write-Host "$zipFilePath" -ForegroundColor White
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Pause-Script
}

# === Run-CleanupScriptData ===
function Run-CleanupScriptData {
    if (-not (Test-Path $ExportDir)) {
        Write-Host "`n⚠️ Export folder not found at $ExportDir" -ForegroundColor Yellow
        Pause-Script
        return
    }

    $files = Get-ChildItem -Path $ExportDir -Recurse -Force -File
    if ($files.Count -eq 0) {
        Write-Host "`nℹ️ No export files to delete." -ForegroundColor Cyan
        Pause-Script
        return
    }

    Write-Host "`n⚠️ The following files will be deleted from $ExportDir:" -ForegroundColor Red
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

# === Menu Loop ===
function Show-ValidationMenuA {
    do {
        Write-Host "`n======= 🧰 Validation Tool A Menu =======" -ForegroundColor Cyan
        Write-Host "[1] Office Validation"
        Write-Host "[2] Driver Validation"
        Write-Host "[3] Roaming Profile Applications"
        Write-Host "[4] Zip and Email Results"
        Write-Host "[5] Cleanup Export Folder"
        Write-Host "[Q] Quit"
        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileApplications }
            "4" { Run-ZipAndEmailResults }
            "5" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-ValidationMenuA