# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Network Tools                         ║
# ║ Version: 1.1 | 2025-07-21                                  ║
# ║ Includes TLS 1.0, ValidateSMB, Npcap Installer             ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-TLS10Check {
    Show-Header "TLS 1.0 Cipher Scan (Port 3389)"
    $target = Read-Host "Enter target IP address"
    $nmap = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"

    if (-not (Test-Path $nmap)) {
        Write-Host "Nmap not found at expected location: $nmap" -ForegroundColor Red
        Pause-Script; return
    }

    $logPath = Get-ExportPath -BaseName "TLS10Scan-$target" -Ext "txt"

    try {
        & $nmap --script ssl-enum-ciphers -p 3389 $target | Out-File $logPath -Encoding UTF8
        Write-ExportPath $logPath
    } catch {
        Write-Host "TLS scan failed: $_" -ForegroundColor Red
    }

    Pause-Script
}

function Run-ValidateSMB {
    Show-Header "ValidateSMB Tool"

    $exe = "$env:ProgramFiles(x86)\CyberCNSAgent\validatesmb.exe"
    if (-not (Test-Path $exe)) {
        Write-Host "Downloading ValidateSMB..." -ForegroundColor Yellow
        try {
            Invoke-WebRequest -Uri "https://betadev.mycybercns.com/agents/validatesmb/validatesmb.exe" -OutFile $exe
        } catch {
            Write-Host "❌ Download failed." -ForegroundColor Red
            Pause-Script; return
        }
    }

    $domain = Read-Host "Enter domain (FQDN)"
    $user   = Read-Host "Enter SMB username"
    $pass   = Read-Host "Enter SMB password"
    $target = Read-Host "Enter target SMB IP"

    $cmd = "`"$exe`" validatesmb $domain $user $pass $target"
    try {
        $outputRaw = Invoke-Expression $cmd
    } catch {
        $outputRaw = "❌ ValidateSMB execution failed: $_"
    }

    $shareTest = Test-Path "\\$target\admin$"

    $results = @(
        [PSCustomObject]@{ Check = "ValidateSMB Output"; Result = ($outputRaw -join "`n") },
        [PSCustomObject]@{ Check = "Admin$ Access Test"; Result = if ($shareTest) { "Accessible" } else { "Failed" } }
    )

    Export-Data -Object $results -BaseName "ValidateSMB"
    Pause-Script
}

function Run-NpcapInstaller {
    Show-Header "Install Npcap (1.79)"
    $url = "https://npcap.com/dist/npcap-1.79.exe"
    $installer = "$env:TEMP\npcap-1.79.exe"

    try {
        Invoke-WebRequest -Uri $url -OutFile $installer
        Start-Process -FilePath $installer -ArgumentList "/S" -Wait -NoNewWindow
        Write-Host "✅ Npcap installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "❌ Npcap install failed: $_" -ForegroundColor Red
    }

    Pause-Script
}

function Show-NetworkToolsMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Network Tools         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] TLS 1.0 Check (Port 3389)",
        " [2] ValidateSMB Tool",
        " [3] Install Npcap",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-TLS10Check }
        "2" { Run-ValidateSMB }
        "3" { Run-NpcapInstaller }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-NetworkToolsMenu
}

Show-NetworkToolsMenu
