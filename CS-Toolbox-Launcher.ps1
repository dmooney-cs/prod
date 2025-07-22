# ╔════════════════════════════════════════════════════════════════╗
# ║  🧰 ConnectSecure Technicians Toolbox – Launcher                ║
# ║  📦 Version: Beta1 – 2025-07-22                                 ║
# ║  🚀 Loads 7 modular tools designed to assist with ConnectSecure ║
# ║     software, troubleshooting, validation, and automation      ║
# ╚════════════════════════════════════════════════════════════════╝

function Show-LauncherMenu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          🧰 ConnectSecure Technicians Toolbox       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "🧩 [1] Validation Tool A        → Office, Drivers, Roaming, Extensions"
    Write-Host "🧩 [2] Validation Tool B        → Patches, VC++ Runtime Scanner"
    Write-Host "🧩 [3] Validation Tool C        → SSL Ciphers, OSQuery Extensions"
    Write-Host "🔧 [4] Agent Maintenance        → Status, Clear Jobs, Set SMB"
    Write-Host "📥 [5] Agent Install Tool       → Download, Install, Configure Agent"
    Write-Host "📡 [6] Network Tools            → TLS 1.0 Scan, Nmap, Validate SMB"
    Write-Host "📂 [7] Active Directory Tools   → Users, Groups, OUs, GPOs"
    Write-Host ""
    Write-Host "❌ [Q] Quit" -ForegroundColor Red
    Write-Host ""
}

function Load-Tool {
    param([string]$url)
    try {
        irm $url | iex
    } catch {
        Write-Host "`n⚠️  Failed to load tool from:" -ForegroundColor Red
        Write-Host $url -ForegroundColor Yellow
        Pause
    }
}

do {
    Show-LauncherMenu
    $choice = Read-Host "Select an option"

    switch ($choice) {
        '1' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20A.ps1" }
        '2' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20B.ps1" }
        '3' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20C.ps1" }
        '4' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Maintenance.ps1" }
        '5' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Install-Tool.ps1" }
        '6' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Network-Tools.ps1" }
        '7' { Load-Tool "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1" }
        'Q' { return }
        default {
            Write-Host "Invalid option. Try again." -ForegroundColor Yellow
            Pause
        }
    }
} while ($true)
