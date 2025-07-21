# ╔════════════════════════════════════════════════════════════════╗
# ║  🧰 ConnectSecure Technicians Toolbox – Launcher                ║
# ║  🚀 Loads 5 modular tools from GitHub links (live)              ║
# ╚════════════════════════════════════════════════════════════════╝

function Show-LauncherMenu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          🧰 ConnectSecure Technicians Toolbox       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "🧩 [1] Validation Collection Tool     → Microsoft, Drivers, Patches"
    Write-Host "🌐 [2] Network Testing Tool           → SMB, TLS 1.0, Nmap Scans"
    Write-Host "🛠  [3] Agent Management Tool         → Install, Status, Clear Jobs"
    Write-Host "👥 [4] Active Directory Tool          → Users, Groups, OUs, GPOs"
    Write-Host "🧪 [5] Utilities Tool                 → Dependency Walker, Cleanup"
    Write-Host ""
    Write-Host "❌ [Q] Exit" -ForegroundColor Yellow
}

function Launch-Tool {
    param([string]$toolUrl)
    try {
        Write-Host "`n▶ Loading tool from: $toolUrl" -ForegroundColor Gray
        irm $toolUrl | iex
    } catch {
        Write-Host "❌ Failed to load script: $_" -ForegroundColor Red
    }
    Write-Host "`nPress any key to return to launcher..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

do {
    Show-LauncherMenu()
    $choice = Read-Host "`nSelect a tool"

    switch ($choice) {
        "1" {
            Launch-Tool -toolUrl "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection.ps1"
        }
        "2" {
            Launch-Tool -toolUrl "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Network.ps1"
        }
        "3" {
            Launch-Tool -toolUrl "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Agent.ps1"
        }
        "4" {
            Launch-Tool -toolUrl "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1"
        }
        "5" {
            Launch-Tool -toolUrl "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Tools-Utilities.ps1"
        }
        "Q" { break }
        default {
            Write-Host "Invalid selection. Try again." -ForegroundColor Yellow
        }
    }
} while ($true)
