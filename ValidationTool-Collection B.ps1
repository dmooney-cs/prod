# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool B             ║
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

function Run-BrowserExtensionAudit {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "$ExportDir\BrowserExtensions_$timestamp`_$hostname.csv"
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
    $csvPath = "$ExportDir\OSquery-browserext-$timestamp-$hostname.csv"
    $jsonPath = "$ExportDir\OSquery-browserext-$timestamp-$hostname.json"

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
    $csvPath = "$ExportDir\TLS443Scan-$hostname-$timestamp.csv"
    $logPath = "$ExportDir\SSLCipherScanLog-$timestamp.csv"

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

    $getHotfixFile = "$ExportDir\GetHotFix_$timestamp`_$hostname.csv"
    $wmicFile = "$ExportDir\WMIC_Hotfix_$timestamp`_$hostname.csv"
    $osqFile = "$ExportDir\OSQuery_Hotfix_$timestamp`_$hostname.csv"

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

function Run-AllValidationsB {
    Run-BrowserExtensionAudit
    Run-OSQueryBrowserExtensions
    Run-SSLCipherValidation
    Run-WindowsPatchDetails
}

function Show-ValidationMenuB {
    do {
        Write-Host "`n======= 🧰 Validation Tool B Menu =======" -ForegroundColor Cyan
        Write-Host "[1] Browser Extension Audit"
        Write-Host "[2] OSQuery Browser Extensions"
        Write-Host "[3] SSL Cipher Validation"
        Write-Host "[4] Windows Patch Details"
        Write-Host "[5] Run All Validations (B)"
        Write-Host "[6] Zip and Email Results"
        Write-Host "[7] Cleanup Export Folder"
        Write-Host "[Q] Quit"
        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-BrowserExtensionAudit }
            "2" { Run-OSQueryBrowserExtensions }
            "3" { Run-SSLCipherValidation }
            "4" { Run-WindowsPatchDetails }
            "5" { Run-AllValidationsB }
            "6" { Run-ZipAndEmailResults }
            "7" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-ValidationMenuB