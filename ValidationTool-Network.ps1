# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Network Tool                     ║
# ╚════════════════════════════════════════════════════╝

function Run-ValidateSMBStatus {
    Write-Host "▶ Checking SMB Version Status..." -ForegroundColor Cyan
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $csvPath = "C:\Script-Export\SMB_Version_Report_$timestamp.csv"
    $results = @()

    $keys = @{
        "SMB 1.0" = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\SMB1"
        "SMB 2.0" = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\SMB2"
    }

    foreach ($version in $keys.Keys) {
        $enabled = "Not Present"
        $disableKey = "N/A"
        try {
            $val = Get-ItemPropertyValue -Path $keys[$version] -Name "(default)" -ErrorAction Stop
            $enabled = if ($val -eq 1) { "Enabled" } else { "Disabled" }
            $disableKey = $keys[$version]
        } catch {
            $enabled = "Missing"
        }
        $results += [PSCustomObject]@{
            SMBVersion     = $version
            Status         = $enabled
            DisableKey     = $disableKey
            DetectionKey   = $keys[$version]
        }
    }

    $results += [PSCustomObject]@{
        SMBVersion   = "SMB 3.0"
        Status       = "Enabled (default)"
        DisableKey   = "Cannot be disabled via registry"
        DetectionKey = "Built-in"
    }

    $results | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "✔ SMB status exported to: $csvPath" -ForegroundColor Green
}

function Run-TLS10Scan {
    Write-Host "▶ Running TLS 1.0 Nmap Scan..." -ForegroundColor Cyan
    $target = Read-Host "Enter the target IP address to scan"
    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapPath)) {
        Write-Host "❌ Nmap not found. Please ensure CyberCNS Agent is installed." -ForegroundColor Red
        return
    }

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outPath = "C:\Script-Export\TLS10Scan-$target-$ts-$hostname.txt"
    $cmd = "`"$nmapPath`" --script ssl-enum-ciphers -p 3389 $target -oN `"$outPath`""
    Invoke-Expression $cmd
    Write-Host "✔ TLS 1.0 scan saved to: $outPath" -ForegroundColor Green
}

function Run-FullNmapScan {
    Write-Host "▶ Running Full Nmap Scan..." -ForegroundColor Cyan
    $target = Read-Host "Enter target IP or range"
    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapPath)) {
        Write-Host "❌ Nmap not found. Please ensure CyberCNS Agent is installed." -ForegroundColor Red
        return
    }

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $output = "C:\Script-Output"
    if (-not (Test-Path $output)) { New-Item -Path $output -ItemType Directory | Out-Null }
    $outFile = "$output\FullNmapScan-$hostname-$ts.txt"

    $scanCmd = "`"$nmapPath`" --privileged -sV -T3 --min-parallelism 100 --max-parallelism 255 --top-ports 3000 -Pn $target -oN `"$outFile`""
    Invoke-Expression $scanCmd

    Write-Host "✔ Full scan results saved to: $outFile" -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportDir\NetworkExport_$hostname_$timestamp.zip"

    if (Test-Path $exportDir) {
        Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile -Force
        Write-Host "✔ ZIP file created: $zipFile" -ForegroundColor Green

        $to = Read-Host "Enter recipient email"
        $mailto = "mailto:$to?subject=Network%20Scan%20Results&body=Results%20attached.%20ZIP:%20$zipFile"
        Start-Process $mailto
    } else {
        Write-Host "❌ No export folder found." -ForegroundColor Yellow
    }
}

function Run-CleanupScriptData {
    Write-Host "🧹 Cleaning up export folder..." -ForegroundColor Red
    Remove-Item -Path "C:\Script-Export\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "✔ Cleanup complete."
}

function Show-NetworkMenu {
    do {
        Write-Host "`n========= 🌐 Network Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Check SMB Versions"
        Write-Host "[2] TLS 1.0 Scan on Port 3389"
        Write-Host "[3] Full Nmap Scan"
        Write-Host "[4] Zip and Email Results"
        Write-Host "[5] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-ValidateSMBStatus }
            "2" { Run-TLS10Scan }
            "3" { Run-FullNmapScan }
            "4" { Run-ZipAndEmailResults }
            "5" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-NetworkMenu
