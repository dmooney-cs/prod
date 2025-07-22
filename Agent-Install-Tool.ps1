# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Installer Utility               ║
# ║ Version: 1.8 | Exact match to one-liner install behavior    ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

$TempDir = "C:\Script-Temp"
if (-not (Test-Path $TempDir)) {
    New-Item -Path $TempDir -ItemType Directory | Out-Null
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
        Write-Host "✅ Agent downloaded successfully." -ForegroundColor Green

        $installCmd = "`"$installer`" -c $company -e $tenant -j $secret -i"
        Write-Host "`n🚀 Installing with command:" -ForegroundColor Cyan
        Write-Host "$installCmd" -ForegroundColor Yellow
        Start-Sleep -Seconds 2

        Start-Process -FilePath $installer -ArgumentList "-c $company -e $tenant -j $secret -i" -Wait
        Write-Host "`n✅ Agent installed successfully." -ForegroundColor Green
        $result = "Success"
    } catch {
        Write-Host "`n❌ Installation failed: $_" -ForegroundColor Red
        $result = "Failed"
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
