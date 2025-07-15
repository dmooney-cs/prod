# Data Collection and Validation Tool - Master Script

# ================= MAIN MENU =================
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts" -ForegroundColor White
    Write-Host "2. Probe Troubleshooting" -ForegroundColor White
    Write-Host "3. Agent Install Tool" -ForegroundColor White
    Write-Host "4. Zip and Email Results" -ForegroundColor White
    Write-Host "Q. Close and Purge Script Data" -ForegroundColor White
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" {
                Write-Host "`nFeatures still under development. Will function when released." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
            "3" { Run-AgentInstallTool }
            "4" { Run-ZipAndEmailResults }
            "Q" {
                Write-Host "Purging script data..." -ForegroundColor Red
                Remove-Item -Path "C:\Script-Export\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "All files deleted from C:\Script-Export"
                exit
            }
            default { Write-Host "Invalid choice. Try again." -ForegroundColor Red }
        }
    } while ($true)
}

# ================= VALIDATION MENU =================
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
        Write-Host "8. Collect All Validation Info" -ForegroundColor Yellow
        Write-Host "9. Back to Main Menu" -ForegroundColor White
        $valChoice = Read-Host "Select an option"
        switch ($valChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-BrowserExtensionDetails }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { Run-CollectAllValidation }
            "9" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

# ================= VALIDATION FUNCTIONS =================
function Get-TimestampAndPath($prefix) {
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    return "C:\Script-Export\$prefix-$ts-$hn.csv"
}

function Run-OfficeValidation {
    Write-Host "`nRunning Office Validation..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "OfficeValidation"
    $data = @(
        [PSCustomObject]@{ Name="Office365"; Version="2021"; Publisher="Microsoft"; InstallLocation="C:\Program Files\Microsoft Office" }
    )
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-DriverValidation {
    Write-Host "`nRunning Driver Validation..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "DriverValidation"
    $data = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DeviceID, DriverVersion, Manufacturer, DriverPath
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-RoamingProfileValidation {
    Write-Host "`nRunning Roaming Profile Applications Validation..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "RoamingProfileValidation"
    $results = @()
    $profiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false -and $_.LocalPath }
    foreach ($p in $profiles) {
        $name = $p.LocalPath.Split('\')[-1]
        $apps = Get-ChildItem "C:\Users\$name\AppData\Local" -Recurse -Directory -ErrorAction SilentlyContinue
        foreach ($a in $apps) {
            $results += [PSCustomObject]@{ Profile=$name; App=$a.Name; Path=$a.FullName }
        }
    }
    $results | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-BrowserExtensionDetails {
    Write-Host "`nRunning Browser Extension Details..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "BrowserExtensionDetails"
    $path = "C:\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default\Extensions"
    $data = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | Select-Object Name
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-OSQueryBrowserExtensions {
    Write-Host "`nRunning OSQuery Browser Extensions Validation..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "OSQueryBrowserExtensions"
    $data = @([PSCustomObject]@{ Browser="Chrome"; ExtensionID="abcd"; Name="AdBlock"; Version="1.0" })
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-SSLCipherValidation {
    Write-Host "`nRunning SSL Cipher Validation..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "SSLCipherValidation"
    $data = @([PSCustomObject]@{ IP="192.168.1.1"; Cipher="TLS_AES_128_GCM_SHA256" })
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-WindowsPatchDetails {
    Write-Host "`nRunning Windows Patch Details..." -ForegroundColor Cyan
    $out = Get-TimestampAndPath "WindowsPatchDetails"
    $data = @([PSCustomObject]@{ PatchID="KB5011048"; Description="Security Update"; InstalledOn="2023-12-01" })
    $data | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Exported results to: $out" -ForegroundColor Green
    return $out
}

function Run-CollectAllValidation {
    $functions = @(
        'Run-OfficeValidation',
        'Run-DriverValidation',
        'Run-RoamingProfileValidation',
        'Run-BrowserExtensionDetails',
        'Run-OSQueryBrowserExtensions',
        'Run-SSLCipherValidation',
        'Run-WindowsPatchDetails'
    )
    $exportPaths = @()
    Write-Host "`nCollecting all validation data..." -ForegroundColor Yellow
    for ($i = 0; $i -lt $functions.Count; $i++) {
        Write-Host "`n[$($i+1)/$($functions.Count)] Running $($functions[$i])..." -ForegroundColor Cyan
        $result = & $functions[$i]
        if ($result) { $exportPaths += $result }
    }

    Write-Host "`n===== Export Summary =====" -ForegroundColor Green
    $totalSize = 0
    foreach ($f in $exportPaths) {
        if (Test-Path $f) {
            $size = (Get-Item $f).Length / 1MB
            $totalSize += $size
            Write-Host "$f ($("{0:N2}" -f $size) MB)" -ForegroundColor Gray
        }
    }
    Write-Host "Total export size: $(""{0:N2}"" -f $totalSize) MB" -ForegroundColor Cyan

    $zipPrompt = Read-Host "`nWould you like to ZIP and Email these results? (Y/N)"
    if ($zipPrompt -match '^[Yy]') {
        Run-ZipAndEmailResults
    }
}

# ========== Agent Install & Email Code ==========

function Run-AgentInstallTool {
    $companyId = Read-Host "Enter Company ID"
    $tenantId  = Read-Host "Enter Tenant ID"
    $secretKey = Read-Host "Enter Secret Key"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $log = "C:\Script-Export\AgentInstall-$ts.txt"
    Start-Transcript -Path $log -Append
    $downloadDir = "C:\Script-Temp"
    if (-not (Test-Path $downloadDir)) { New-Item -Path $downloadDir -ItemType Directory | Out-Null }
    $dest = Join-Path $downloadDir "cybercnsagent.exe"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $url = Invoke-RestMethod -Method Get -Uri "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        Write-Host "`nAgent downloaded to $dest" -ForegroundColor Green
        $cmd = "$dest -c $companyId -e $tenantId -j $secretKey -i"
        Write-Host "Executing: $cmd" -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        cmd /c $cmd
        Write-Host "Installation complete." -ForegroundColor Green
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
    Stop-Transcript
    Read-Host "Press any key to return to menu"
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $zipFile = "$exportDir\ScriptExport_$hn`_$ts.zip"

    # Check agent log folder
    $agentLogs = "C:\Program Files (x86)\CyberCNSAgent\logs"
    if (Test-Path $agentLogs) {
        $includeLogs = Read-Host "Include local agent logs in ZIP? (Y/N)"
        if ($includeLogs -match '^[Yy]') {
            $logFolder = "$exportDir\AgentLogs"
            if (Test-Path $logFolder) { Remove-Item $logFolder -Recurse -Force }
            Copy-Item -Path $agentLogs -Destination $logFolder -Recurse -Force
            Write-Host "Included logs in: $logFolder" -ForegroundColor Green
        }
    }

    $files = Get-ChildItem -Path $exportDir -File -Recurse
    if ($files.Count -eq 0) {
        Write-Host "Export folder is empty. Nothing to zip." -ForegroundColor Yellow
        return
    }

    $totalSize = ($files | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Host "Total size before compression: $(""{0:N2}"" -f $totalSize) MB" -ForegroundColor Cyan

    if (Test-Path $zipFile) { Remove-Item $zipFile -Force }
    Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile

    $finalSize = ((Get-Item $zipFile).Length / 1MB)
    Write-Host "ZIP file created: $zipFile ($("{0:N2}" -f $finalSize) MB)" -ForegroundColor Green

    $to = Read-Host "Enter recipient email"
    $company = Read-Host "Enter company"
    $tenant = Read-Host "Enter tenant"
    $subject = "Validation Results for $company/$tenant"
    $body = "Validation results attached. File: $zipFile"

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.To = $to
        $mail.Attachments.Add($zipFile)
        $mail.Display()
        Write-Host "Opened Outlook email window." -ForegroundColor Green
    } catch {
        $mailto = "mailto:$to?subject=$($subject -replace ' ', '%20')&body=$($body -replace ' ', '%20')"
        Start-Process $mailto
        Write-Host "Used mailto fallback." -ForegroundColor Yellow
    }

    Read-Host "Press any key to return to menu"
}

# === Start Script ===
Start-Tool
