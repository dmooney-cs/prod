# Data Collection and Validation Tool - Master Script

# Function to show the main menu
function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts" -ForegroundColor White
    Write-Host "2. Probe Troubleshooting" -ForegroundColor White
    Write-Host "3. Agent Install Tool" -ForegroundColor White
    Write-Host "4. Zip and Email Results" -ForegroundColor White
    Write-Host "Q. Close and Purge Script Data" -ForegroundColor White
}

# Function to show the Validation Scripts menu
function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Microsoft Office Validation" -ForegroundColor White
        Write-Host "2. Driver Validation" -ForegroundColor White
        Write-Host "3. Roaming Profile Applications" -ForegroundColor White
        Write-Host "4. Browser Extension Details" -ForegroundColor White
        Write-Host "5. OSQuery Browser Extensions" -ForegroundColor White
        Write-Host "6. SSL Cipher Validation" -ForegroundColor White
        Write-Host "7. Windows Patch Details" -ForegroundColor White
        Write-Host "8. Back to Main Menu" -ForegroundColor White
        $valChoice = Read-Host "Select an option"
        
        switch ($valChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-BrowserExtensionDetails }
            "5" { Run-OSQueryBrowserExtensions }
            "6" { Run-SSLCipherValidation }
            "7" { Run-WindowsPatchDetails }
            "8" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

# === Agent Install Tool ===
function Show-AgentInstallToolMenu {
    Write-Host "`nRunning the Agent Install Tool..." -ForegroundColor Cyan
    Run-Install
}

# --- Agent Install Code from provided example ---
function Run-Install {
    $companyId = Read-Host "Enter Company ID"
    $tenantId  = Read-Host "Enter Tenant ID"
    $secretKey = Read-Host "Enter Secret Key"

    Write-Host "`nUsing TLS 1.2 for secure agent link download..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Host "Fetching agent download URL from ConnectSecure API..." -ForegroundColor Cyan
    try {
        $source = Invoke-RestMethod -Method "Get" -Uri "https://configuration.myconnectsecure.com/api/v4/configuration/agentlink?ostype=windows"
    } catch {
        Write-Host "Failed to retrieve download URL: $_" -ForegroundColor Red
        return
    }

    $downloadDir = "C:\Script-Temp"
    if (-not (Test-Path $downloadDir)) {
        New-Item -Path $downloadDir -ItemType Directory | Out-Null
    }

    $destination = Join-Path $downloadDir "cybercnsagent.exe"

    Write-Host "Downloading agent to $destination" -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $source -OutFile $destination -UseBasicParsing
        Write-Host "Agent downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download agent: $_" -ForegroundColor Red
        return
    }

    $installCmd = "$destination -c $companyId -e $tenantId -j $secretKey -i"
    Write-Host "`nExecuting: $installCmd" -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    cmd /c $installCmd

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Agent installation completed successfully." -ForegroundColor Green
    } else {
        Write-Host "Agent installation failed (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
    }

    # Final prompt to exit after installation
    Read-Host -Prompt "`nPress any key to exit"
}

# --- Zip and Email Results ---
function Show-ZipAndEmailMenu {
    Run-ZipAndEmailResults
}

function Run-ZipAndEmailResults {
    # Set up folder and output paths
    $exportFolder = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFilePath = "$exportFolder\ScriptExport_${hostname}_$timestamp.zip"

    # Ensure export folder exists
    if (-not (Test-Path $exportFolder)) {
        Write-Host "Folder '$exportFolder' not found. Exiting..." -ForegroundColor Red
        exit
    }

    # Get all files recursively
    $allFiles = Get-ChildItem -Path $exportFolder -Recurse -File
    if ($allFiles.Count -eq 0) {
        Write-Host "No files found in '$exportFolder'. Exiting..." -ForegroundColor Yellow
        exit
    }

    # Display contents
    Write-Host ""
    Write-Host "=== Contents of $exportFolder ===" -ForegroundColor Cyan
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
    Compress-Archive -Path "$exportFolder\*" -DestinationPath $zipFilePath

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
    Read-Host "Press ENTER to exit..."
}

# Main function to start the tool
function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        
        switch ($choice.ToUpper()) {
            "1" { Run-ValidationScripts }
            "2" { 
                Write-Host "`nFeatures still under development. Will function when released." -ForegroundColor Yellow
                Start-Sleep -Seconds 2  # Wait for 2 seconds to show the message
            }
            "3" { Show-AgentInstallToolMenu }
            "4" { Show-ZipAndEmailMenu }  # Added Zip and Email option
            "Q" { 
                Write-Host "Purging script data..." -ForegroundColor Red
                Remove-Item -Path "C:\Script-Export\*" -Recurse -Force
                Write-Host "All files deleted from C:\Script-Export"
                exit 
            }
            default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            }
        }
    } while ($true)
}

# Start the main tool
Start-Tool
