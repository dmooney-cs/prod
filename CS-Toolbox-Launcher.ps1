<#
CS-Toolbox-Launcher.ps1
Version: Beta1
Date: 2025-07-21

Description:
Loads 5 modular tools designed to assist in troubleshooting, installing, and utilizing ConnectSecure software and services.
#>

# Define GitHub raw URLs for each sub-tool
$scriptUrls = @{
    "Tools-Utilities"     = "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Tools-Utilities.ps1"
    "ValidationTool-AD"   = "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/ValidationTool-AD.ps1"
    "Validation-Scripts"  = "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Validation-Scripts.ps1"
    "Agent-Maintenance"   = "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Agent-Maintenance.ps1"
    "Network-Tools"       = "https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Network-Tools.ps1"
}

# Load all scripts into memory
foreach ($name in $scriptUrls.Keys) {
    try {
        Write-Host "Loading $name module..." -ForegroundColor Cyan
        Invoke-RestMethod -Uri $scriptUrls[$name] | Invoke-Expression
    }
    catch {
        Write-Warning "Failed to load $name from $($scriptUrls[$name])"
    }
}

function Show-LauncherMenu {
    Clear-Host
    Write-Host "==============================================" -ForegroundColor DarkCyan
    Write-Host "       ConnectSecure Toolbox Launcher" -ForegroundColor Cyan
    Write-Host "             Version: Beta1 (2025-07-21)" -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "1. Run Validation Scripts" -ForegroundColor Green
    Write-Host "2. Run Agent Maintenance" -ForegroundColor Green
    Write-Host "3. Run Network Tools" -ForegroundColor Green
    Write-Host "4. Run AD Validation Tools" -ForegroundColor Green
    Write-Host "5. Run Utilities" -ForegroundColor Green
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host ""

    $choice = Read-Host "Enter your choice"
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationMenu }
        "2" { Run-AgentMaintenance }
        "3" { Run-NetworkMenu }
        "4" { Run-ADValidationMenu }
        "5" { Run-ToolsUtilities }
        "Q" { return }
        default {
            Write-Warning "Invalid choice. Please try again."
            Start-Sleep -Seconds 1.5
            Show-LauncherMenu
        }
    }
}

# Launch the main menu
Show-LauncherMenu
