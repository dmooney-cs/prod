# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Agent Installer Utility               ║
# ║ Version: 1.0 | 2025-07-21                                  ║
# ║ Downloads & installs the CyberCNS Agent                    ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder
$TempDir = "C:\Script-Temp"; if (-not (Test-Path $TempDir)) { New-Item $TempDir -ItemType Directory | Out-Null }

function Run-AgentInstaller {
    Show-Header "CyberCNS Agent Install"

    $company = Read-Host "Enter Company ID"
    $tenant  = Read-Host "Enter Tenant ID"
    $secret  = Read-Host "Enter Secret Key"

    $agentUrl = "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    $outfile  = "$TempDir\cybercnsagent.exe"

    try {
        Write-Host "Downloading agent installer..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $agentUrl -OutFile $outfile -UseBasicParsing
        Write-Host "Agent downloaded to: $outfile" -ForegroundColor Green
    } catch {
        Write-Host "Download failed: $_" -ForegroundColor Red
        Pause-Script; return
    }

    $cmd = "`"$outfile`" -c $company -e $tenant -j $secret -i"
    Write-Host "Executing: $cmd" -ForegroundColor Cyan
    Start-Sleep -Seconds 5

    try {
        Start-Process -FilePath $outfile -ArgumentList "-c $company -e $tenant -j $secret -i" -Wait
        Write-Host "`n✅ Agent installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "`n❌ Installation failed: $_" -ForegroundColor Red
    }

    Pause-Script
}

Run-AgentInstaller
<Recovered Agent Installer>
