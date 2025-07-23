# ╔════════════════════════════════════════════════════════════════╗
# ║  🧰 ConnectSecure Technicians Toolbox – Launcher                ║
# ║  🚀 Loads 5 modular tools from GitHub links (live)              ║
# ║  Version: Beta1 | 2025-07-23                                    ║
# ╚════════════════════════════════════════════════════════════════╝

function Show-LauncherMenu {
    Clear-Host
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          🧰 ConnectSecure Technicians Toolbox       ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "🧩 [1] Validation Tool A         → Office, Drivers, Roaming, Extensions"
    Write-Host "🧩 [2] Validation Tool B         → Patches, VC++ Runtime Scanner"
    Write-Host "🧩 [3] Validation Tool C         → SSL Ciphers, OSQuery Extensions"
    Write-Host "🛠  [4] Agent Maintenance         → Status, Clear Jobs, Set SMB"
    Write-Host "🧪 [5] Agent Install Tool        → Download, Install, Configure Agent"
    Write-Host "🌐 [6] Network Tools             → TLS 1.0 Scan, Nmap, Validate SMB"
    Write-Host "📁 [7] Active Directory Tools    → Users, Groups, OUs, GPOs"
    Write-Host ""
    Write-Host "[Q] Quit"
    Write-Host ""
}

do {
    Show-LauncherMenu
    $choice = Read-Host "Enter your choice"

    switch ($choice.ToUpper()) {
        '1' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20A.ps1")
        }
        '2' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20B.ps1")
        }
        '3' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-Collection%20C.ps1")
        }
        '4' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Maintenance.ps1")
        }
        '5' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Install-Tool.ps1")
        }
        '6' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Network-Tools.ps1")
        }
        '7' {
            iex (Invoke-RestMethod -Uri "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1")
        }
        'Q' {
            Write-Host "`nExiting Toolbox. Goodbye!" -ForegroundColor Yellow
            break
        }
        default {
            Write-Host "`nInvalid selection. Please choose a valid option." -ForegroundColor Red
        }
    }

    if ($choice -ne 'Q') {
        Write-Host "`nPress any key to return to the menu..." -ForegroundColor DarkGray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }

} while ($choice -ne 'Q')
