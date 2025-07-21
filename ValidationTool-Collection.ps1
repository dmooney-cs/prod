# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Collection Tool               ║
# ║ Version: Beta1 | 2025-07-21                        ║
# ╚════════════════════════════════════════════════════╝

function Run-OfficeValidation {
    Write-Host "▶ Running Office Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\OfficeValidation-$ts-$hn.csv"
    $data = @(
        [PSCustomObject]@{ Name="Office365"; Version="2021"; Publisher="Microsoft"; InstallLocation="C:\Program Files\Microsoft Office" }
    )
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-DriverValidation {
    Write-Host "▶ Running Driver Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\DriverValidation-$ts-$hn.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DeviceID, DriverVersion, Manufacturer, DriverPath
    $drivers | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-RoamingProfileValidation {
    Write-Host "▶ Running Roaming Profile Applications Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\RoamingProfileValidation-$ts-$hn.csv"
    $data = @()
    $profiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath -ne $null }
    foreach ($p in $profiles) {
        $name = $p.LocalPath.Split('\')[-1]
        $apps = Get-ChildItem "C:\Users\$name\AppData\Local" -Recurse -Directory -ErrorAction SilentlyContinue
        foreach ($a in $apps) {
            $data += [PSCustomObject]@{ Profile=$name; App=$a.Name; Path=$a.FullName }
        }
    }
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-BrowserExtensionDetails {
    Write-Host "▶ Running Browser Extension Details..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\BrowserExtensionDetails-$ts-$hn.csv"
    $browserPaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Extensions",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Extensions",
        "$env:APPDATA\Mozilla\Firefox\Profiles"
    )
    $data = @()
    foreach ($path in $browserPaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $data += [PSCustomObject]@{ Browser=$path.Split('\')[4]; Extension=$_.Name; Path=$_.FullName }
            }
        }
    }
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-OSQueryBrowserExtensions {
    Write-Host "▶ Running OSQuery Browser Extensions Validation..." -ForegroundColor Cyan
    $osq = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    if (-not (Test-Path $osq)) {
        Write-Host "OSQuery not found." -ForegroundColor Red
        return
    }
    $query = "SELECT * FROM chrome_extensions;"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $csv = "C:\Script-Export\OSQueryBrowserExtensions-$ts-$hn.csv"
    $output = & $osq --json "$query" | ConvertFrom-Json
    $output | Export-Csv -Path $csv -NoTypeInformation
    Write-Host "Exported results to: $csv" -ForegroundColor Green
}

function Run-SSLCipherValidation {
    Write-Host "▶ Running SSL Cipher Validation..." -ForegroundColor Cyan
    $nmap = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\TLS443Scan-$hn-$ts.csv"
    if (-not (Test-Path $nmap)) {
        Write-Host "Nmap not found." -ForegroundColor Red
        return
    }
    $cmd = "$nmap --script ssl-enum-ciphers -p 443 127.0.0.1 -oN $out"
    Invoke-Expression $cmd
    Write-Host "Nmap scan saved to $out" -ForegroundColor Green
}

function Run-WindowsPatchDetails {
    Write-Host "▶ Running Windows Patch Details..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\WindowsPatchDetails-$ts-$hn.csv"
    $patches = Get-HotFix | Select-Object Description, HotFixID, InstalledOn
    $patches | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-VCRuntimeDependencyCheck {
    Write-Host "▶ Running VC++ Dependency Check..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\VCRuntimeCheck-$ts-$hn.csv"
    $data = @()
    $locations = @("$env:SystemRoot\System32", "$env:SystemRoot\SysWOW64")
    foreach ($loc in $locations) {
        Get-ChildItem -Path $loc -Include *.dll -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.Name -match "vcruntime|msvcp|msvcr") {
                $data += [PSCustomObject]@{ File = $_.Name; Path = $_.FullName }
            }
        }
    }
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportDir\ScriptExport_$hostname_$timestamp.zip"

    if (-not (Test-Path $exportDir)) {
        Write-Host "No export folder found." -ForegroundColor Yellow
        return
    }

    Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile -Force
    Write-Host "ZIP file created: $zipFile" -ForegroundColor Green

    $to = Read-Host "Enter recipient email"
    $encodedPath = $zipFile -replace '\\', '/' -replace ':', '%3A' -replace ' ', '%20'
    $subject = [uri]::EscapeDataString("Tool Export Results")
    $body = [uri]::EscapeDataString("Please manually attach the ZIP located at: $encodedPath")
    $mailto = "mailto:$to?subject=$subject&body=$body"

    # Absolute fix using cmd /c start with quoted mailto
    $launch = "start `"`"$mailto`"`""
    cmd.exe /c $launch
}

function Run-CleanupScriptData {
    Write-Host "🧹 Cleaning up export folder..." -ForegroundColor Red
    Remove-Item -Path "C:\Script-Export\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "Cleanup complete."
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
        Write-Host "[8] VC++ Runtime & Dependency Check"
        Write-Host "[9] Zip and Email Results"
        Write-Host "[10] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-BrowserExtensionDetails }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { Run-VCRuntimeDependencyCheck }
            "9" { Run-ZipAndEmailResults }
            "10" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-CollectionMenu
