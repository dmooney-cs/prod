# 🔧 CS Toolbox – Shared Functions v3.12

function Show-Header {
    param ([string]$Title)
    $width = 50
    $padded = "║   $Title".PadRight($width - 1) + "║"
    Clear-Host
    Write-Host ""
    Write-Host ("╔" + ("═" * ($width - 2)) + "╗") -ForegroundColor Cyan
    Write-Host $padded -ForegroundColor Cyan
    Write-Host ("╚" + ("═" * ($width - 2)) + "╝") -ForegroundColor Cyan
    Write-Host ""
}

function Pause-Script {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor DarkGray
    try { [void][System.Console]::ReadKey($true) } catch { Read-Host "Press ENTER to continue" }
}

function Ensure-ExportFolder {
    $Global:ExportFolder = "C:\Script-Export"
    if (-not (Test-Path $ExportFolder)) {
        New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
    }
}

function Get-ExportPath {
    param ($BaseName, $Ext = "csv")
    $hn = $env:COMPUTERNAME
    $time = Get-Date -Format "yyyyMMdd_HHmmss"
    return "$ExportFolder\$BaseName-$time-$hn.$Ext"
}

function Export-Data {
    param (
        [Parameter(Mandatory)] $Object,
        [Parameter(Mandatory)] [string] $BaseName,
        [string] $Ext = "csv"
    )
    if (-not $Object -or $Object.Count -eq 0) {
        Write-Host "No data to export for $BaseName." -ForegroundColor Yellow
        return
    }
    $path = Get-ExportPath -BaseName $BaseName -Ext $Ext
    switch ($Ext) {
        "csv"  { $Object | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $path }
        "json" { $Object | ConvertTo-Json -Depth 5 | Out-File $path -Encoding UTF8 }
        "txt"  { $Object | Out-File $path -Encoding UTF8 }
        default {
            Write-Host "Unsupported export format: $Ext" -ForegroundColor Red
            return
        }
    }
    Write-ExportPath $path
}

function Write-ExportPath {
    param ($Path)
    Write-Host "`nExport saved to: $Path" -ForegroundColor Green
}

function Show-FolderContents {
    param ([string]$Folder)
    if (Test-Path $Folder) {
        Write-Host "📂 Export folder: $Folder" -ForegroundColor Gray
    } else {
        Write-Host "Folder not found: $Folder" -ForegroundColor Yellow
    }
}

function Invoke-ZipAndEmailResults {
    Show-Header "Zip and Email Export Results"
    Ensure-ExportFolder

    $company = Read-Host "Enter Company Name"
    $tenant  = Read-Host "Enter Tenant Name"

    $logsPath = "C:\Program Files (x86)\CyberCNSAgent\logs"
    $logFolder = ""
    $failedFiles = @()

    if (Test-Path $logsPath) {
        $logAnswer = Read-Host "Include local agent logs from '$logsPath' in export? (Y/N)"
        if ($logAnswer -eq "Y") {
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $hostname = $env:COMPUTERNAME
            $logFolder = Join-Path $ExportFolder "AgentLogs_${stamp}_$hostname"
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null

            Get-ChildItem -Path $logsPath -File | ForEach-Object {
                $dest = Join-Path $logFolder $_.Name
                try {
                    Get-Content $_.FullName -ErrorAction Stop | Out-File -FilePath $dest -Force -Encoding UTF8
                } catch {
                    $failedFiles += $_.Name
                }
            }

            if ($failedFiles.Count -gt 0) {
                Write-Host "`n⚠️ Could not read/copy the following files due to lock or binary format:" -ForegroundColor Yellow
                $failedFiles | ForEach-Object { Write-Host " - $_" -ForegroundColor DarkGray }
            }

            $copied = (Get-ChildItem -Path $logFolder -File).Count
            Write-Host "✅ Copied $copied log files to: $logFolder" -ForegroundColor Green
        }
    }

    $zipName = "ExportResults_{0}_{1}.zip" -f $env:COMPUTERNAME, (Get-Date -Format "yyyyMMdd_HHmmss")
    $zipPath = Join-Path $ExportFolder $zipName
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    Show-FolderContents -Folder $ExportFolder

    Write-Host "`n🔍 Files and folders to be zipped:" -ForegroundColor DarkGray
    Get-ChildItem -Path $ExportFolder | ForEach-Object {
        $size = if ($_.PSIsContainer) { "[DIR]" } else { "$([math]::Round($_.Length / 1KB, 1)) KB" }
        Write-Host " - $($_.Name) $size" -ForegroundColor Gray
    }

    try {
        Compress-Archive -Path "$ExportFolder\*" -DestinationPath $zipPath -Force
        $zipSize = (Get-Item $zipPath).Length
        Write-Host "`n✅ Main ZIP created: $zipSize bytes" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to create export zip: $_" -ForegroundColor Red
        return
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.To = "support@connectsecure.com"
        $mail.Subject = "CS Toolbox Results - $env:COMPUTERNAME | $company / $tenant"
        $mail.Body = "Attached is the ZIP file containing validation output from this system."
        $mail.Attachments.Add($zipPath)
        $mail.Display()
        Write-Host "📧 Outlook message created with attachment." -ForegroundColor Green
    } catch {
        $mailto = "mailto:support@connectsecure.com?subject=CS Toolbox Results - $env:COMPUTERNAME ($company / $tenant)&body=Please attach ZIP file manually: $zipName"
        Start-Process $mailto
        Write-Host "⚠️ Outlook not available. Launching default mail client..." -ForegroundColor Yellow
    }

    Write-Host "`nZIP path: $zipPath" -ForegroundColor Cyan
    Pause-Script
}

function Invoke-CleanupExportFolder {
    Show-Header "Clean Up Export Folder"

    if (!(Test-Path $ExportFolder)) {
        Write-Host "No export folder found to clean." -ForegroundColor Yellow
        Pause-Script
        return
    }

    $items = Get-ChildItem -Path $ExportFolder -Recurse -Force
    if (-not $items) {
        Write-Host "Folder already empty." -ForegroundColor Yellow
        Pause-Script
        return
    }

    Write-Host "⚠️  Files to be deleted: $($items.Count)" -ForegroundColor Red
    $confirm = Read-Host "Are you sure you want to delete all contents of $ExportFolder? (Y/N)"
    if ($confirm -ne "Y") {
        Write-Host "Aborted cleanup." -ForegroundColor Yellow
        Pause-Script
        return
    }

    try {
        $items | Remove-Item -Force -Recurse
        Write-Host "✅ Export folder cleaned." -ForegroundColor Green
    } catch {
        Write-Host "❌ Cleanup failed: $_" -ForegroundColor Red
    }

    Pause-Script
}
