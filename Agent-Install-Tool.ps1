# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Installer Tool                  ║
# ║ Version: 1.3 | Detects version, includes install menu       ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder
$TempDir = "C:\Script-Temp"
if (-not (Test-Path $TempDir)) {
    New-Item -Path $TempDir -ItemType Directory | Out-Null
}

$installer = "$TempDir\cybercnsagent.exe"
$agentUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
$detectedVersion = "Unknown"

# Download and detect version before menu
try {
    Invoke-WebRequest -Uri $agentUrl -OutFile $installer -UseBasicParsing
    if (Test-Path $installer) {
        $versionInfo = (Get-Item $installer).VersionInfo
        $detectedVersion = $versionInfo.ProductVersion
    }
} catch {
    Write-Host "⚠️ Could not download agent or detect version." -ForegroundColor Yellow
}

function Run-AgentInstaller {
    Show-Header "CyberCNS Agent - INSTALL"

    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $secret  = Read-Host "Enter Secret Key"

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $logFile = "$ExportFolder\AgentInstall-Log-$timestamp-$hostname.txt"
    $summaryPath = "$ExportFolder\AgentInstall-Summary-$timestamp-$hostname.csv"
    $version = $detectedVersion
    $result = "Not Run"

    Start-Transcript -Path $logFile -Force

    if (Test-Path $installer) {
        Write-Host "📦 Detected agent version: $version" -ForegroundColor Cyan
        $installCmd = "`"$installer`" -c $company -e $tenant -j $secret -i"
        Write-Host "`n📦 Installing agent with command:" -ForegroundColor Cyan
        Write-Host "$installCmd" -ForegroundColor Yellow
        Start-Sleep -Seconds 3

        try {
            Start-Process -FilePath $installer -ArgumentList "-c $company -e $tenant -j $secret -i" -Wait
            Write-Host "`n✅ Agent installed successfully." -ForegroundColor Green
            $result = "Success"
        } catch {
            Write-Host "`n❌ Agent install failed." -ForegroundColor Red
            $result = "Install failed"
        }
    } else {
        Write-Host "❌ Installer not found. Cannot proceed." -ForegroundColor Red
        $result = "Download failed"
    }

    $summary = [PSCustomObject]@{
        Timestamp    = (Get-Date)
        Hostname     = $hostname
        CompanyID    = $company
        TenantID     = $tenant
        AgentVersion = $version
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
    Write-Host "Please upload and link 'Agent-Reinstall.ps1' when ready." -ForegroundColor DarkGray
    Pause-Script
}

function Run-AgentUninstall {
    Show-Header "CyberCNS Agent - UNINSTALL"
    Write-Host "⚠️ This option will be implemented in a separate script." -ForegroundColor Yellow
    Write-Host "Please upload and link 'Agent-Uninstall.ps1' when ready." -ForegroundColor DarkGray
    Pause-Script
}

function Show-AgentInstallerMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Agent Installer Menu         ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Detected Agent Version: $detectedVersion" -ForegroundColor Magenta
    Write-Host ""

    $menu = @(
        " [1] Install CyberCNS Agent",
        " [2] Reinstall Agent (external)",
        " [3] Uninstall Agent (external)",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-AgentInstaller }
        "2" { Run-AgentReinstall }
        "3" { Run-AgentUninstall }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-AgentInstallerMenu
}

Show-AgentInstallerMenu
