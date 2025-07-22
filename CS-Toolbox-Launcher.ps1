# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Launcher                               ║
# ║ Version: Beta3 | 2025-07-21                                 ║
# ║ Loads 9 modular tools with visible subtools                ║
# ╚═════════════════════════════════════════════════════════════╝

function Show-LauncherMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║ 🧰 CS Tech Toolbox – Launcher Menu                ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    Write-Host " [1] Validation Tool A"
    Write-Host "     ├─ Office Detection"
    Write-Host "     ├─ Driver Audit"
    Write-Host "     ├─ Roaming Profile Apps"
    Write-Host "     └─ Browser Extension Details"

    Write-Host " [2] Validation Tool B"
    Write-Host "     ├─ VC++ Runtime + Binary Scan"
    Write-Host "     └─ Windows Patch Validation (WMIC, OSQuery, HotFix)"

    Write-Host " [3] Validation Tool C"
    Write-Host "     ├─ OSQuery Browser Extensions"
    Write-Host "     └─ SSL Cipher Validation (Nmap 443)"

    Write-Host " [4] Active Directory Collection"
    Write-Host "     ├─ AD Users"
    Write-Host "     ├─ AD Groups"
    Write-Host "     ├─ AD Computers"
    Write-Host "     └─ AD OUs + GPO Links"

    Write-Host " [5] Network Tools"
    Write-Host "     ├─ TLS 1.0 Check (Port 3389)"
    Write-Host "     ├─ ValidateSMB Tool"
    Write-Host "     └─ Npcap Installer"

    Write-Host " [6] Agent Maintenance"
    Write-Host "     ├─ Check Agent Status"
    Write-Host "     ├─ Clear Pending Jobs"
    Write-Host "     ├─ Set SMB Settings"
    Write-Host "     └─ Check SMB Version"

    Write-Host " [7] Agent Installer Utility"
    Write-Host "     └─ Download and Install CyberCNS Agent"

    Write-Host " [8] System Info A"
    Write-Host "     ├─ Firewall Status"
    Write-Host "     ├─ Microsoft Defender Status"
    Write-Host "     └─ Disk Space + SMART Health"

    Write-Host " [9] System Info B"
    Write-Host "     ├─ Pending Reboot Status"
    Write-Host "     ├─ Event Log Error Summary (72h)"
    Write-Host "     └─ Startup / Autostart Audit"

    Write-Host " [Q] Quit"
    Write-Host ""

    $choice = Read-Host "Select an option"
    switch ($choice) {
        "1" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20A.ps1 | iex }
        "2" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20B.ps1 | iex }
        "3" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20C.ps1 | iex }
        "4" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1 | iex }
        "5" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Network-Tools.ps1 | iex }
        "6" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Maintenance.ps1 | iex }
        "7" { irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Install-Tool.ps1 | iex }
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

