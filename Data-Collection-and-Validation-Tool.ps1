
# Data-Collection-and-Validation-Tool.ps1 - Restored Working Version

function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
    Write-Host "====================================================="
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Microsoft Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Roaming Profile Applications"
        Write-Host "4. Browser Extension Details"
        Write-Host "5. SSL Cipher Validation"
        Write-Host "6. Windows Patch Details"
        Write-Host "7. Back to Main Menu"
        $valChoice = Read-Host "Select an option"

        switch ($valChoice) {
            "1" { Run-OfficeValidation }
            "2" { Run-DriverValidation }
            "3" { Run-RoamingProfileValidation }
            "4" { Run-OSQueryBrowserExtensions }
            "5" { Run-SSLCipherValidation }
            "6" { Run-WindowsPatchDetails }
            "7" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-ZipAndEmail {
    $exportFolder = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportFolder\SystemExport_$timestamp`_$hostname.zip"

    if (-not (Test-Path $exportFolder)) {
        Write-Host "Export folder does not exist at $exportFolder" -ForegroundColor Red
        return
    }

    $files = Get-ChildItem -Path $exportFolder -Recurse -File
    if ($files.Count -eq 0) {
        Write-Host "No files found to zip in $exportFolder" -ForegroundColor Yellow
        return
    }

    Write-Host "`nThe following files will be zipped:" -ForegroundColor Cyan
    $files | ForEach-Object { Write-Host "• $($_.FullName)" -ForegroundColor Gray }

    $confirmZip = Read-Host "`nWould you like to create a ZIP archive of these files? (Y/N)"
    if ($confirmZip -notin @("Y", "y")) {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }

    if (Test-Path $zipFile) { Remove-Item -Path $zipFile -Force -ErrorAction SilentlyContinue }
    Compress-Archive -Path "$exportFolder\*" -DestinationPath $zipFile -Force

    Write-Host "`n✅ ZIP archive created:" -ForegroundColor Green
    Write-Host $zipFile -ForegroundColor Cyan

    $nextStep = Read-Host "`nWould you like to send this ZIP via email or return to the main menu? (Email/Menu)"
    if ($nextStep -match "Email") {
        $recipient = Read-Host "Enter the email address to send the ZIP file to"

        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0)
            $Mail.Subject = "System Validation Export - $hostname"
            $Mail.To = $recipient
            $Mail.Body = "Attached is the exported validation data from $hostname."
            $Mail.Attachments.Add($zipFile)
            $Mail.Display()
            Write-Host "Email draft opened in Outlook." -ForegroundColor Green
        } catch {
            Write-Host "Outlook not available. Attempting to use default email client." -ForegroundColor Yellow
            Start-Process "explorer.exe" -ArgumentList "/select,$zipFile"
            $mailto = "mailto:$recipient?subject=System Validation Export - $hostname"
            Start-Process $mailto
        }
    } else {
        Write-Host "Returning to main menu..." -ForegroundColor Cyan
    }
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" -ForegroundColor Cyan }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" -ForegroundColor Cyan }
        "4" { Run-ZipAndEmail }
        "Q" { exit }
        default { Write-Host "Invalid option." -ForegroundColor Red }
    }
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
        if ($choice -ne "Q") {
            Write-Host "`nPress Enter to return to menu..."
            Read-Host | Out-Null
            Clear-Host
        }
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
