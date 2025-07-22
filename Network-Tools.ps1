# ╔════════════════════════════════════════════════════════════╗
# ║ 🌐 CS Tech Toolbox – Network Tools                         ║
# ║ Version: N.3 – TLS 1.0, ValidateSMB, Npcap Install         ║
# ╚════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-TLS10Scan {
    Clear-Host
    Write-Host "`n=== TLS 1.0 Cipher Scan (Port 3389) ===`n" -ForegroundColor Cyan

    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapPath)) {
        Write-Host "❌ Nmap not found at $nmapPath" -ForegroundColor Red
        Pause-Script
        return
    }

    $ip = Read-Host "Enter target IP address"
    if (-not $ip) {
        Write-Host "Cancelled." -ForegroundColor Yellow
        Pause-Script
        return
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "C:\Script-Export\TLS10Scan-$ip-$timestamp-$hostname.txt"

    & $nmapPath --script ssl-enum-ciphers -p 3389 $ip | Tee-Object -FilePath $outputFile

    Write-Host "`n✔ Scan complete. Results saved to:" -ForegroundColor Green
    Write-Host $outputFile -ForegroundColor Yellow
    Pause-Script
}

function Run-ValidateSMB {
    Clear-Host
    Write-Host "`n=== ValidateSMB Tool ===`n" -ForegroundColor Cyan

    $toolPath = "C:\Program Files (x86)\CyberCNSAgent\ValidateSMB.exe"
    if (-not (Test-Path $toolPath)) {
        Write-Host "❌ ValidateSMB.exe not found." -ForegroundColor Red
        Pause-Script
        return
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputFile = "C:\Script-Export\ValidateSMB-$timestamp-$hostname.txt"

    & $toolPath | Tee-Object -FilePath $outputFile

    Write-Host "`n✔ SMB results saved to:" -ForegroundColor Green
    Write-Host $outputFile -ForegroundColor Yellow
    Pause-Script
}

function Run-InstallNpcap {
    Clear-Host
    Write-Host "`n=== Npcap Installer ===`n" -ForegroundColor Cyan

    $url = "https://npcap.com/dist/npcap-1.79.exe"
    $installer = "C:\Script-Temp\npcap-1.79.exe"

    if (-not (Test-Path "C:\Script-Temp")) {
        New-Item -Path "C:\Script-Temp" -ItemType Directory | Out-Null
    }

    Write-Host "⬇ Downloading Npcap..."
    Invoke-WebRequest -Uri $url -OutFile $installer

    if (Test-Path $installer) {
        Write-Host "⚙ Installing Npcap silently..."
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait

        Write-Host "`n✔ Npcap installation attempted. You may verify manually." -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to download Npcap." -ForegroundColor Red
    }

    Pause-Script
}

function Show-NetworkMenu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         🌐 CS Toolbox – Network Tools Menu         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [1] TLS 1.0 Check (Port 3389)"
    Write-Host " [2] ValidateSMB Tool"
    Write-Host " [3] Install Npcap"
    Write-Host ""
    Write-Host " [Z] Zip and Email Results"
    Write-Host " [C] Cleanup Export Folder"
    Write-Host " [Q] Quit"
    Write-Host ""
}

do {
    Show-NetworkMenu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Run-TLS10Scan }
        '2' { Run-ValidateSMB }
        '3' { Run-InstallNpcap }
        'Z' { Run-ZipAndEmailResults }
        'C' { Run-CleanupExportFolder }
        'Q' { return }
        default {
            Write-Host "Invalid option. Try again." -ForegroundColor Yellow
            Pause-Script
        }
    }
} while ($true)
