# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Installer Utility               ║
# ║ Version: 1.0 | 2025-07-21                                  ║
# ║ Includes version check, install, logging, ZIP, cleanup     ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder
$TempDir = "C:\Script-Temp"; if (-not (Test-Path $TempDir)) { New-Item $TempDir -ItemType Directory | Out-Null }

function Get-AgentVersion {
    $url = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    $headVer = $getVer = $null

    Write-Host "🔍 Checking for latest agent version..."

    # HEAD check
    Write-Host "➡ Attempting HEAD request to extract version..."
    try {
        $resp = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
        $headVer = $resp.Headers["X-Agent-Version"] ?? ($resp.BaseResponse.ResponseUri.Segments[-1])
        if ($headVer) {
            Write-Host "✅ Found version in headers: $headVer"
        }
    } catch {
        Write-Host "❌ HEAD version check failed." -ForegroundColor Yellow
    }

    # GET check fallback
    Write-Host "➡ Attempting GET request to extract version..."
    try {
        $get = Invoke-WebRequest -Uri $url -Method Get -UseBasicParsing
        $getVer = ($get.BaseResponse.ResponseUri.Segments[-1])
        if ($getVer) {
            Write-Host "✅ Found version in GET redirect: $getVer"
        }
    } catch {
        Write-Host "❌ GET version check failed." -ForegroundColor Yellow
    }

    return ($headVer ?? $getVer ?? "Unknown")
}

function Run-AgentInstaller {
    Show-Header "CyberCNS Agent Installer"

    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $secret  = Read-Host "Enter Secret Key"
    $version = Get-AgentVersion

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $logFile = "$ExportFolder\AgentInstall-Log-$timestamp-$hostname.txt"
    $summaryPath = "$ExportFolder\AgentInstall-Summary-$timestamp-$hostname.csv"
    $installer = "$TempDir\cybercnsagent.exe"
    $agentUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"

    Start-Transcript -Path $logFile -Force

    Write-Host "`n📥 Downloading latest agent..."
    $downloaded = $false
    try {
        Invoke-WebRequest -Uri $agentUrl -OutFile $installer -UseBasicParsing
        $downloaded = $true
        Write-Host "✅ Agent downloaded to: $installer" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to download agent." -ForegroundColor Red
    }

    if ($downloaded -and (Test-Path $installer)) {
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

function Show-AgentInstallerMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Agent Installer      ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Install CyberCNS Agent",
        " [2] Zip and Email Export Folder",
        " [3] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-AgentInstaller }
        "2" { Invoke-ZipAndEmailResults }
        "3" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-AgentInstallerMenu
}

Show-AgentInstallerMenu
