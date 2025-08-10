# =====================================================================
# ConnectSecure Technicians Toolbox - Main Launcher (FromZip)
# Version: 2.2.1 (2025-08-10)
# Notes:
#  - Forces session execution policy (Bypass) to avoid prompts.
#  - Loads Functions-Common.ps1.
#  - Uses -f formatting anywhere a variable is followed by a colon to
#    avoid "$var:" parsing errors in PowerShell.
# =====================================================================

# 0) Force session-only execution policy to avoid "Run this script?" prompts
try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
} catch {
    Write-Host "⚠ Unable to set execution policy for this session. Some scripts may prompt." -ForegroundColor Yellow
}

# 1) Load shared functions first – required for Show-Header and others
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$commonPath = Join-Path $scriptRoot 'Functions-Common.ps1'

if (-not (Test-Path $commonPath)) {
    Write-Host ("❌ ERROR: Functions-Common.ps1 not found in {0}" -f $scriptRoot) -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    return
}

try {
    $code = Get-Content -Path $commonPath -Encoding UTF8 -Raw
    Invoke-Expression $code
} catch {
    Write-Host ("❌ ERROR: Failed to load Functions-Common.ps1: {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    return
}

Ensure-ExportFolder

# 2) Banner
Clear-Host
Show-Header "ConnectSecure Technicians Toolbox"

# 3) Menu
function Show-MainMenu {
    Write-Host ""
    Write-Host " [1] OSQuery Data Collection        - Browser/Apps via osquery"
    Write-Host " [2] Nmap Data Collection           - Port/Service scan via nmap"
    Write-Host " [3] Agent Menu Tool                - Install, Uninstall, Status, Maintenance"
    Write-Host " [4] Active Directory Tools         - Users, Groups, OUs, GPOs"
    Write-Host " [5] System Info A                  - Firewall, Defender, Disk/SMART"
    Write-Host " [6] System Info B                  - Pending Reboot, App Logs, Startup Audit"
    Write-Host " [7] Utilities                      - Running Services, Disk Space"
    Write-Host ""
    Write-Host " [Z] Zip and Email Results          - Compress results for support"
    Write-Host " [C] Cleanup Toolbox Data           - Remove temp/output and self-clean"
    Write-Host " [Q] Quit"
    Write-Host ""
}

function Invoke-MenuAction {
    param([Parameter(Mandatory=$true)][string]$Choice)

    $toolMap = @{
        '1' = 'Osquery-Data-Collection.ps1'
        '2' = 'Nmap-Data-Collection.ps1'
        '3' = 'Agent-Menu-Tool.ps1'
        '4' = 'ValidationTool-AD.ps1'
        '5' = 'SystemInfo-A.ps1'
        '6' = 'SystemInfo-B.ps1'
        '7' = 'Tools-Utilities.ps1'
        'Z' = 'ZipResults'
        'C' = 'Cleanup'
        'Q' = 'Quit'
    }

    $key = $Choice.ToUpperInvariant()
    if (-not $toolMap.ContainsKey($key)) {
        Write-Host "Invalid selection." -ForegroundColor Yellow
        return $false
    }

    switch ($key) {
        'Z' {
            Zip-Results
            Pause-Script "Press any key to return to the menu..."
            return $false
        }
        'C' {
            Invoke-FinalCleanupAndExit
            return $true
        }
        'Q' { return $true }
        default {
            $toolName = $toolMap[$key]
            $toolPath = Join-Path $scriptRoot $toolName
            if (-not (Test-Path $toolPath)) {
                Write-Host ("ERROR launching {0}: File not found." -f $toolName) -ForegroundColor Red
                Pause-Script
                return $false
            }

            # Launch in a new elevated window with ExecutionPolicy Bypass
            $ok = Launch-Tool -Path $toolPath -Elevated:$true -NewWindow:$true
            if (-not $ok) {
                Write-Host ("ERROR launching {0}." -f $toolName) -ForegroundColor Red
                Pause-Script
            }
            return $false
        }
    }
}

# 4) Loop
while ($true) {
    Show-Header "ConnectSecure Technicians Toolbox"
    Show-MainMenu
    $choice = Read-Host "Enter your choice"
    if ([string]::IsNullOrWhiteSpace($choice)) {
        Write-Host "Please enter a selection." -ForegroundColor Yellow
        Start-Sleep -Milliseconds 700
        continue
    }
    $quit = Invoke-MenuAction -Choice $choice
    if ($quit) { break }
}

Write-Host "Goodbye." -ForegroundColor Cyan
