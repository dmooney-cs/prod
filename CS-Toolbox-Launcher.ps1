# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Launcher                               ║
# ║ Version: Beta3 | 2025-07-21                                 ║
# ║ Loads 9 modular tools with visible subtools                ║
# ╚═════════════════════════════════════════════════════════════╝

function Show-LauncherMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔═════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║ 🧰 CS Tech Toolbox – Launcher Menu                          ║" -ForegroundColor Cyan
    Write-Host "║ Version: Beta3 | 2025-07-21                                 ║" -ForegroundColor Cyan
    Write-Host "║ Loads 9 modular tools with visible subtools                ║" -ForegroundColor Cyan
    Write-Host "╚═════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    Write-Host " 🧩 [1] Validation Tool A        – Office, Drivers, Roaming Apps, Extensions"
    Write-Host " 🧪 [2] Validation Tool B        – VC++ Scan, Windows Patches (WMIC/OSQ/HotFix)"
    Write-Host " 🔐 [3] Validation Tool C        – OSQuery Extensions, SSL Cipher (Nmap 443)"
    Write-Host " 🏢 [4] Active Directory Tools   – Users, Groups, Computers, OUs, GPO Links"
    Write-Host " 🌐 [5] Network Tools            – TLS 1.0 (3389), ValidateSMB, Npcap Installer"
    Write-Host " 🛠️  [6] Agent Maintenance        – Status, Clear Jobs, Set/Check SMB"
    Write-Host " 🚀 [7] Agent Installer Utility  – Install, Uninstall, Zip/Email, Cleanup"
    Write-Host " 💽 [8] System Info A            – Firewall, Defender, Disk Space & SMART"
    Write-Host " 📋 [9] System Info B            – Reboot Status, Logs, Startup Items"
    Write-Host " ❌ [Q] Quit"
    Write-Host ""

    $choice = Read-Host "Select an option"
    switch ($choice) {
        "1" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-CollectionA-Fixed.ps1 | iex }
        "2" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20B.ps1 | iex }
        "3" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20C.ps1 | iex }
        "4" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1 | iex }
        "5" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Network-Tools.ps1 | iex }
        "6" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Maintenance.ps1 | iex }
        "7" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Menu-Tool.ps1 | iex }
        "8" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/SystemInfo-A.ps1 | iex }
        "9" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/SystemInfo-B.ps1 | iex }
        "Q" {
            Write-Host "`nGoodbye!" -ForegroundColor Green
            return
        }
        default {
            Write-Host "`nInvalid option." -ForegroundColor Red
        }
    }

    Write-Host "`nPress any key to return to the launcher..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
    Show-LauncherMenu
}

Show-LauncherMenu
