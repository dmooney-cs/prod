# Data Collection and Validation Tool - Full Script (Final Adjustments)

# Create export directory
$ExportDir = "C:\Script-Export"
if (-not (Test-Path $ExportDir)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

# Function to pause for user input (handles environment issues)
function Pause-Script {
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    try {
        [void][System.Console]::ReadKey($true)  # Attempt to use the Console.ReadKey method
    } catch {
        Read-Host "Press Enter to continue..."  # Fallback for environments that cannot use ReadKey
    }
}

# Main Menu Display
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
    $selection = Read-Host "Select an option"
    switch ($selection) {
        '1' { Show-ValidationMenu }
        '2' { Show-AgentMaintenanceMenu }
        '3' { Show-ProbeTroubleshootingMenu }
        '4' { Run-ZipAndEmailResults }
        'Q' { Cleanup-And-Exit }
        default { Show-MainMenu }
    }
}

# Validation Scripts Menu
function Show-ValidationMenu {
    Clear-Host
    Write-Host "======== Validation Scripts Menu ========" -ForegroundColor Cyan
    Write-Host "1. Microsoft Office Validation"
    Write-Host "2. Driver Validation"
    Write-Host "3. Roaming Profile Applications"
    Write-Host "4. Browser Extension Details"
    Write-Host "5. OSQuery Browser Extensions"
    Write-Host "6. SSL Cipher Validation"
    Write-Host "7. Windows Patch Details"
    Write-Host "8. Active Directory Collection"
    Write-Host "9. Back to Main Menu"
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Run-OfficeValidation; Show-ValidationMenu }
        '2' { Run-DriverValidation; Show-ValidationMenu }
        '3' { Run-RoamingProfileValidation; Show-ValidationMenu }
        '4' { Run-BrowserExtensionDetails; Show-ValidationMenu }
        '5' { Run-OSQueryBrowserExtensions; Show-ValidationMenu }
        '6' { Run-SSLCipherValidation; Show-ValidationMenu }
        '7' { Run-WindowsPatchDetails; Show-ValidationMenu }
        '8' { Run-ActiveDirectoryValidation; Show-ValidationMenu }
        '9' { Show-MainMenu }
        default { Show-ValidationMenu }
    }
}

# Function to run the Zip and Email Results operation
function Run-ZipAndEmailResults {
    $ExportFolder = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$ExportFolder\ScriptExport_${hostname}_$timestamp.zip"

    # Ensure export folder exists
    if (-not (Test-Path $ExportFolder)) {
        Write-Host "Folder '$ExportFolder' not found. Exiting..." -ForegroundColor Red
        exit
    }

    # Get all files recursively
    $allFiles = Get-ChildItem -Path $ExportFolder -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$ExportFolder'. Exiting..." -ForegroundColor Yellow
        exit
    }

    # Display contents
    Write-Host ""
    Write-Host "=== Contents of $ExportFolder ===" -ForegroundColor Cyan
    $allFiles | ForEach-Object { Write-Host $_.FullName }

    # Calculate total size before compression
    $totalSizeMB = [math]::Round(($allFiles | Measure-Object Length -Sum).Sum / 1MB, 2)
    Write-Host ""
    Write-Host "Total size before compression: $totalSizeMB MB" -ForegroundColor Yellow

    # Prompt to ZIP
    $zipChoice = Read-Host "`nWould you like to zip the folder contents? (Y/N)"
    if ($zipChoice -notin @('Y','y')) {
        Write-Host "Zipping skipped. Exiting..." -ForegroundColor DarkGray
        exit
    }

    # Remove old zip if exists
    if (Test-Path $zipFilePath) { Remove-Item $zipFilePath -Force }

    # Create ZIP
    Compress-Archive -Path "$ExportFolder\*" -DestinationPath $zipFilePath

    # Get ZIP size
    $zipSizeMB = [math]::Round((Get-Item $zipFilePath).Length / 1MB, 2)

    # Show summary
    Write-Host ""
    Write-Host "=== ZIP Summary ===" -ForegroundColor Green
    Write-Host "ZIP File: $zipFilePath"
    Write-Host "Size after compression: $zipSizeMB MB" -ForegroundColor Yellow

    # Prompt to email
    $emailChoice = Read-Host "`nWould you like to email the ZIP file? (Y/N)"
    if ($emailChoice -in @('Y','y')) {
        $recipient = Read-Host "Enter recipient email address"
        $subject = "Script Export from $hostname"
        $body = "Attached is the export ZIP file from $hostname.`nZIP Path: $zipFilePath"

        Write-Host "`nChecking for Outlook..." -ForegroundColor Cyan
        $Outlook = $null
        $OutlookWasRunning = $false

        # Attempt to connect to Outlook (running or start new)
        try {
            $Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $OutlookWasRunning = $true
        } catch {
            try {
                $Outlook = New-Object -ComObject Outlook.Application
                $OutlookWasRunning = $false
            } catch {
                $Outlook = $null
            }
        }

        if ($Outlook) {
            try {
                # Ensure MAPI session is open
                $namespace = $Outlook.GetNamespace("MAPI")
                $namespace.Logon($null, $null, $false, $true)

                # Create and populate mail item
                $Mail = $Outlook.CreateItem(0)
                $Mail.To = $recipient
                $Mail.Subject = $subject
                $Mail.Body = $body

                if (Test-Path $zipFilePath) {
                    $Mail.Attachments.Add($zipFilePath)
                } else {
                    Write-Host "❌ ZIP file not found at $zipFilePath" -ForegroundColor Red
                }

                $Mail.Display()
                Write-Host "`n✅ Outlook draft email opened with ZIP attached." -ForegroundColor Green
            } catch {
                Write-Host "❌ Outlook COM error: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "`n⚠️ Outlook not available. Falling back to default mail client..." -ForegroundColor Yellow
            $mailto = "mailto:$recipient`?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
            Start-Process $mailto
            Write-Host ""
            Write-Host "Please manually attach this file:" -ForegroundColor Cyan
            Write-Host "$zipFilePath" -ForegroundColor White
        }
    } else {
        Write-Host "Email skipped." -ForegroundColor DarkGray
    }

    Write-Host ""
    Pause-Script  # Pauses before returning to menu
}

# Cleanup & Exit Logic
function Cleanup-And-Exit {
    Write-Host "`nCleaning up all Script Data..." -ForegroundColor Yellow
    $pathsToDelete = @("C:\Script-Export", "C:\Script-Temp")
    foreach ($path in $pathsToDelete) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Host "Script data purged successfully. Exiting..." -ForegroundColor Green
    Start-Sleep -Seconds 5
    exit
}

# Launch the menu
Show-MainMenu
