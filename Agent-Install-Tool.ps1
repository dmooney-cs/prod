# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Installer Utility               ║
# ║ Version: 2.2 | Service check, status report, uninstall ask ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

$TempDir = "C:\Script-Temp"
if (-not (Test-Path $TempDir)) {
    New-Item -Path $TempDir -ItemType Directory | Out-Null
}

function Run-AgentInstaller {
    Show-Header "CyberCNS Agent - INSTALL"

    # Step 0: Check for service presence and status
    Write-Host "`n[Checking ConnectSecure Agent service status...]" -ForegroundColor Gray
    $svc1 = Get-Service -Name "CyberCNSAgent" -ErrorAction SilentlyContinue
    $svc2 = Get-Service -Name "CyberCNSAgentMonitor" -ErrorAction SilentlyContinue

    $svc1Status = if ($svc1) {
        "✅ Installed | " + ($(if ($svc1.Status -eq 'Running') {"✅ Running"} else {"❌ Not Running"}))
    } else {
        "❌ Not Installed"
    }

    $svc2Status = if ($svc2) {
        "✅ Installed | " + ($(if ($svc2.Status -eq 'Running') {"✅ Running"} else {"❌ Not Running"}))
    } else {
        "❌ Not Installed"
    }

    Write-Host "CyberCNSAgent:        $svc1Status"
    Write-Host "CyberCNSAgentMonitor: $svc2Status"

    if ($svc1 -and $svc2 -and $svc1.Status -ne 'Running' -and $svc2.Status -ne 'Running') {
        $startPrompt = Read-Host "`nBoth services are installed but not running.`nWould you like to start them now? (Y/N)"
        if ($startPrompt -match '^[Yy]$') {
            try {
                Start-Service -Name CyberCNSAgent, CyberCNSAgentMonitor -ErrorAction Stop
                Write-Host "✅ Services started successfully." -ForegroundColor Green
            } catch {
                Write-Host "❌ Failed to start services: $_" -ForegroundColor Red
                Pause-Script; return
            }
        }
    }

    if ($svc1 -and $svc1.Status -eq 'Running') {
        $reinstallPrompt = Read-Host "`nConnectSecure Agent is already running.`nWould you like to uninstall it before reinstalling? (Y/N)"
        if ($reinstallPrompt -match '^[Yy]$') {
            Run-AgentUninstall
            return
        }
    }

    # Proceed with install
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $logFile = "$ExportFolder\AgentInstall-Log-$timestamp-$hostname.txt"
    $summaryPath = "$ExportFolder\AgentInstall-Summary-$timestamp-$hostname.csv"
    $installer = "$TempDir\cybercnsagent.exe"
    $agentApiUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    $result = "Not Run"

    Start-Transcript -Path $logFile -Force

    try {
        Write-Host "`n🔒 Setting TLS 1.2..." -ForegroundColor DarkGray
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        Write-Host "`n🔍 Resolving agent download link..."
        $source = Invoke-RestMethod -Method Get -Uri $agentApiUrl
        Write-Host "➡ Download URL: $source" -ForegroundColor Gray

        Write-Host "`n📥 Downloading agent to: $installer"
        Invoke-WebRequest -Uri $source -OutFile $installer -UseBasicParsing
        $size = (Get-Item $installer).Length

        Write-Host "📦 Downloaded file size: $size bytes"
        if ($size -lt 100000) {
            Write-Host "❌ Downloaded file too small — likely invalid." -ForegroundColor Red
            throw "Agent download integrity check failed."
        }

        Write-Host "✅ Agent downloaded and validated." -ForegroundColor Green
    } catch {
        Write-Host "`n❌ Agent download failed: $_" -ForegroundColor Red
        Stop-Transcript
        Pause-Script
        return
    }

    # Prompt for install inputs only if download succeeded
    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $secret  = Read-Host "Enter Secret Key"

    try {
        $installCmd = "`"$installer`" -c $company -e $tenant -j $secret -i"
        Write-Host "`n🚀 Running inline install command:" -ForegroundColor Cyan
        Write-Host "$installCmd" -ForegroundColor Yellow
        Start-Sleep -Seconds 2

        & "$installer" -c $company -e $tenant -j $secret -i

        Write-Host "`n✅ Agent executed inline successfully." -ForegroundColor Green
        $result = "Success"
    } catch {
        Write-Host "`n❌ Agent execution failed: $_" -ForegroundColor Red
        $result = "Execution failed"
    }

    $summary = [PSCustomObject]@{
        Timestamp    = (Get-Date)
        Hostname     = $hostname
        CompanyID    = $company
        TenantID     = $tenant
        Result       = $result
    }

    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    Write-ExportPath $summaryPath

    Stop-Transcript
    Write-ExportPath $logFile

    Pause-Script
}

function Run-AgentReinstall {
    Show-Header "CyberCNS Agent - REINSTALL"
    Write-Host "⚠️ This option will be implemented in a separate script." -ForegroundColor Yellow
    Pause-Script
}

function Run-AgentUninstall {
    Show-Header "CyberCNS Agent - UNINSTALL"
    Write-Host "⚠️ This option will be implemented in a separate script." -ForegroundColor Yellow
    Pause-Script
}

function Show-AgentInstallerMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Agent Installer Menu         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    $menu = @(
        " [1] Install CyberCNS Agent",
        " [2] Reinstall Agent (external)",
        " [3] Uninstall Agent (external)",
        " [4] Zip and Email Export Folder",
        " [5] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-AgentInstaller }
        "2" { Run-AgentReinstall }
        "3" { Run-AgentUninstall }
        "4" { Invoke-ZipAndEmailResults }
        "5" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-AgentInstallerMenu
}

Show-AgentInstallerMenu
