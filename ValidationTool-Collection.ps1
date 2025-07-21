# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Collection Tool               ║
# ║ Version: Beta1 | 2025-07-21                        ║
# ╚════════════════════════════════════════════════════╝

$ExportDir = "C:\Script-Export"
if (-not (Test-Path $ExportDir)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

function Pause-Script {
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    try { [void][System.Console]::ReadKey($true) } catch { Read-Host "Press Enter to continue..." }
}

# ... [All other validation functions omitted here for brevity] ...

function Run-ZipAndEmailResults {
    $ExportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$ExportDir\ScriptExport_${hostname}_$timestamp.zip"

    # Check if user wants to include Agent logs
    $agentLogSource = "C:\Program Files (x86)\CyberCNSAgent\logs"
    $agentLogTarget = "$ExportDir\AgentLogs"
    $includeLogs = Read-Host "Include Local ConnectSecure Agent Logs @ $agentLogSource? (Y/N)"
    if ($includeLogs -in @("Y","y")) {
        if (Test-Path $agentLogSource) {
            if (-not (Test-Path $agentLogTarget)) {
                New-Item -Path $agentLogTarget -ItemType Directory | Out-Null
            }
            Copy-Item -Path "$agentLogSource\*" -Destination $agentLogTarget -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "✔ Agent logs copied to: $agentLogTarget" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Agent log folder not found at: $agentLogSource" -ForegroundColor Yellow
        }
    }

    if (-not (Test-Path $ExportDir)) {
        Write-Host "Folder '$ExportDir' not found." -ForegroundColor Red
        Pause-Script
        return
    }

    $allFiles = Get-ChildItem -Path $ExportDir -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$ExportDir'." -ForegroundColor Yellow
        Pause-Script
        return
    }

    Compress-Archive -Path "$ExportDir\*" -DestinationPath $zipFilePath -Force
    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)
    Write-Host "`nZIP created: $zipFilePath ($zipSizeMB MB)" -ForegroundColor Green

    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $bodyFallback = "Please attach the file located at: $zipFilePath before sending this email."

        $modernOutlookPath = "$env:LOCALAPPDATA\Packages\microsoft.windowscommunicationsapps_8wekyb3d8bbwe"
        $isModernOutlook = Test-Path $modernOutlookPath

        Write-Host "`nChecking for Outlook..." -ForegroundColor Cyan
        if ($isModernOutlook) {
            Write-Host "🧭 New Outlook (Microsoft Store version) detected. COM automation is not supported." -ForegroundColor Yellow
        }

        try {
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            try { $Outlook = New-Object -ComObject Outlook.Application } catch { $Outlook = $null }
        }

        if ($Outlook) {
            try {
                $namespace = $Outlook.GetNamespace("MAPI")
                $namespace.Logon($null, $null, $false, $true)
                $Mail = $Outlook.CreateItem(0)
                $Mail.To = $recipient
                $Mail.Subject = $subject
                $Mail.Body = "Attached is the export ZIP file from $hostname.`nZIP Path: $zipFilePath"
                if (Test-Path $zipFilePath) {
                    $Mail.Attachments.Add($zipFilePath)
                }
                $Mail.Display()
                Write-Host "`n✅ Outlook draft email opened with ZIP attached." -ForegroundColor Green
            } catch {
                Write-Host "❌ Outlook COM error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "`n⚠️ Outlook COM interface not available. Launching default mail client..." -ForegroundColor Yellow
            if ($isModernOutlook) {
                Write-Host "✳️ This appears to be the New Outlook (Microsoft Store version), which cannot auto-attach files." -ForegroundColor Magenta
            }

            $mailto = "mailto:$recipient?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($bodyFallback))"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"$mailto`"" -WindowStyle Hidden
            Write-Host "`nPlease manually attach this file:" -ForegroundColor Cyan
            Write-Host "$zipFilePath" -ForegroundColor White
        }
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Pause-Script
}

function Show-CollectionMenu {
    do {
        Write-Host "`n========= 🧩 Collection Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Office Validation"
        Write-Host "[2] Driver Validation"
        Write-Host "[3] Roaming Profile Applications"
        Write-Host "[4] Browser Extension Details"
        Write-Host "[5] OSQuery Browser Extensions"
        Write-Host "[6] SSL Cipher Validation"
        Write-Host "[7] Windows Patch Details"
        Write-Host "[8] VC++ Runtime & Dependency Check"
        Write-Host "[9] Zip and Email Results"
        Write-Host "[10] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-BrowserExtensionDetails }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { Run-VCRuntimeDependencyCheck }
            "9" { Run-ZipAndEmailResults }
            "10" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-CollectionMenu
