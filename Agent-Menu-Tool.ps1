# ╔════════════════════════════════════════════════════╗
# ║ 🧰 ConnectSecure Agent Management Menu              ║
# ║ Version: 1.0 | Includes install, uninstall, zip     ║
# ╚════════════════════════════════════════════════════╝

function Show-AgentMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║    🧰 ConnectSecure Agent Management Menu           ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [1] Install or Reinstall Agent"
    Write-Host " [2] Uninstall Agent"
    Write-Host " [3] Zip and Email Results"
    Write-Host " [4] Cleanup and Exit"
    Write-Host " [Q] Quit to Launcher"

    $choice = Read-Host "`nSelect an option"
    switch ($choice) {
        "1" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Install-Tool-NoZip.ps1 | iex
        }
        "2" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Uninstall-CyberCNSAgentV4.ps1 | iex
        }
        "3" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-ZipAndEmailResults
        }
        "4" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-CleanupExportFolder
        }
        "Q" {
            return
        }
        default {
            Write-Host "`nInvalid selection." -ForegroundColor Red
            Pause-Script
        }
    }

    Show-AgentMenu
}

function Pause-Script {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor DarkGray
    try { [void][System.Console]::ReadKey($true) } catch { Read-Host "Press ENTER to continue" }
}

Show-AgentMenu
