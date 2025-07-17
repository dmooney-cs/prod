
# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Utilities Tool                   ║
# ╚════════════════════════════════════════════════════╝

function Download-DependencyWalker {
    $tempDir = "C:\Script-Temp"
    $zipUrl = "https://www.dependencywalker.com/depends22_x86.zip"
    $zipPath = "$tempDir\depends22_x86.zip"
    $extractPath = "$tempDir\depends"

    if (-not (Test-Path $tempDir)) {
        New-Item -Path $tempDir -ItemType Directory | Out-Null
    }

    if (-not (Test-Path "$extractPath\depends.exe")) {
        Write-Host "▶ Downloading Dependency Walker..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Host "✔ Dependency Walker downloaded and extracted." -ForegroundColor Green
    } else {
        Write-Host "✔ Dependency Walker already present." -ForegroundColor Yellow
    }
}

function Run-DependencyWalkerScan {
    Download-DependencyWalker
    $exePath = "$env:ProgramFiles (x86)\CyberCNSAgent\nmap\nmap.exe"
    $target = Read-Host "Enter the full path to a .exe or .dll to scan"
    if (-not (Test-Path $target)) {
        Write-Host "❌ File not found: $target" -ForegroundColor Red
        return
    }

    $scanType = Read-Host "Scan for (1) .NET or (2) VC++ Dependencies?"
    $scanFlag = if ($scanType -eq "1") { "/c /pa:$target" } else { "/c /pb:$target" }

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $outPath = "C:\Script-Export\DependencyWalkerScan_${ts}_$hn.txt"

    Write-Host "▶ Running scan..." -ForegroundColor Cyan
    Start-Process -FilePath "C:\Script-Temp\depends\depends.exe" -ArgumentList "$scanFlag > `"$outPath`"" -NoNewWindow -Wait
    Write-Host "✔ Scan results saved to: $outPath" -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportDir\UtilitiesExport_$hostname_$timestamp.zip"

    if (Test-Path $exportDir) {
        Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile -Force
        Write-Host "✔ ZIP file created: $zipFile" -ForegroundColor Green

        $to = Read-Host "Enter recipient email"
        $mailto = "mailto:$to?subject=Utility%20Results&body=Scan%20attached.%20ZIP:%20$zipFile"
        Start-Process $mailto
    } else {
        Write-Host "❌ No export folder found." -ForegroundColor Yellow
    }
}

function Run-CleanupScriptData {
    Write-Host "🧹 Cleaning up export folder..." -ForegroundColor Red
    Remove-Item -Path "C:\Script-Export\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "✔ Cleanup complete."
}

function Show-UtilitiesMenu {
    do {
        Write-Host "`n========= 🧪 Utilities Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Run Dependency Walker Scan"
        Write-Host "[2] Zip and Email Results"
        Write-Host "[3] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-DependencyWalkerScan }
            "2" { Run-ZipAndEmailResults }
            "3" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-UtilitiesMenu
