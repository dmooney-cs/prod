# CS Tech Toolbox – Collection Tool
# Version: Beta1 | 2025-07-21

# Create export directory if needed
$ExportDir = "C:\Script-Export"
if (-not (Test-Path $ExportDir)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

function Pause-Script {
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    try { [void][System.Console]::ReadKey($true) } catch { Read-Host "Press Enter to continue..." }
}

function Run-OfficeValidation {
    Write-Host "▶ Running Office Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\OfficeValidation-$ts-$hn.csv"
    $data = @(
        [PSCustomObject]@{ Name="Office365"; Version="2021"; Publisher="Microsoft"; InstallLocation="C:\Program Files\Microsoft Office" }
    )
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    Pause-Script
}

function Run-DriverValidation {
    Write-Host "▶ Running Driver Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\DriverValidation-$ts-$hn.csv"
    $drivers = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DeviceID, DriverVersion, Manufacturer, DriverPath
    $drivers | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    Pause-Script
}

function Run-RoamingProfileValidation {
    Write-Host "▶ Running Roaming Profile Applications Validation..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\RoamingProfileValidation-$ts-$hn.csv"
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
    Pause-Script
}

function Run-BrowserExtensionDetails {
    Write-Host "▶ Running Browser Extension Details..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\BrowserExtensionDetails-$ts-$hn.csv"
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
    Pause-Script
}

function Run-OSQueryBrowserExtensions {
    Write-Host "▶ Running OSQuery Browser Extensions Validation..." -ForegroundColor Cyan
    $osq = "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe"
    if (-not (Test-Path $osq)) {
        Write-Host "OSQuery not found." -ForegroundColor Red
        Pause-Script
        return
    }
    $query = "SELECT * FROM chrome_extensions;"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $csv = "$ExportDir\OSQueryBrowserExtensions-$ts-$hn.csv"
    $output = & $osq --json "$query" | ConvertFrom-Json
    $output | Export-Csv -Path $csv -NoTypeInformation
    Write-Host "Exported results to: $csv" -ForegroundColor Green
    Pause-Script
}

function Run-SSLCipherValidation {
    Write-Host "▶ Running SSL Cipher Validation..." -ForegroundColor Cyan
    $nmap = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\TLS443Scan-$hn-$ts.csv"
    if (-not (Test-Path $nmap)) {
        Write-Host "Nmap not found." -ForegroundColor Red
        Pause-Script
        return
    }
    $cmd = "$nmap --script ssl-enum-ciphers -p 443 127.0.0.1 -oN $out"
    Invoke-Expression $cmd
    Write-Host "Nmap scan saved to $out" -ForegroundColor Green
    Pause-Script
}

function Run-WindowsPatchDetails {
    Write-Host "▶ Running Windows Patch Details..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\WindowsPatchDetails-$ts-$hn.csv"
    $patches = Get-HotFix | Select-Object Description, HotFixID, InstalledOn
    $patches | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    Pause-Script
}

function Run-VCRuntimeDependencyCheck {
    Write-Host "▶ Running VC++ Dependency Check..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "$ExportDir\VCRuntimeCheck-$ts-$hn.csv"
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
    Pause-Script
}

function Run-ZipAndEmailResults {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$ExportDir\ScriptExport_${hostname}_$timestamp.zip"

    if (-not (Test-Path $ExportDir)) {
        Write-Host "Folder '$ExportDir' not found." -ForegroundColor Red
        Pause-Script
        return
    }

    $allFiles = Get-ChildItem -Path $ExportDir -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$ExportDir'." -ForegroundColor Yellow
        Pause-Script
        return
    }

    Compress-Archive -Path "$ExportDir\*" -DestinationPath $zipFilePath -Force
    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)
    Write-Host "`nZIP created: $zipFilePath ($zipSizeMB MB)" -ForegroundColor Green

    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $body = "Please attach the file located at: $zipFilePath before sending this email."

        Write-Host "`nChecking for Outlook..." -ForegroundColor Cyan
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
                if (Test-Path $zipFilePath) { $Mail.Attachments.Add($zipFilePath) }
                $Mail.Display()
                Write-Host "`n✅ Outlook draft email opened with ZIP attached." -ForegroundColor Green
            } catch {
                Write-Host "❌ Outlook COM error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "`n⚠️ Outlook not available. Launching default mail client..." -ForegroundColor Yellow
            $mailto = "mailto:$recipient?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"$mailto`"" -WindowStyle Hidden
            Write-Host "`nPlease manually attach this file:" -ForegroundColor Cyan
            Write-Host "$zipFilePath" -ForegroundColor White
        }
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }
    Pause-Script
}
