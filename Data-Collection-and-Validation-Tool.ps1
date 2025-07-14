
# Data-Collection-and-Validation-Tool.ps1 - Stable Pre-ZIP Version

function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Exit"
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
            "1" { Write-Host "[Placeholder] Run-OfficeValidation" -ForegroundColor Cyan }
            "2" { Write-Host "[Placeholder] Run-DriverValidation" -ForegroundColor Cyan }
            "3" { Write-Host "[Placeholder] Run-RoamingProfileValidation" -ForegroundColor Cyan }
            "4" { Write-Host "[Placeholder] Run-OSQueryBrowserExtensions" -ForegroundColor Cyan }
            "5" { Write-Host "[Placeholder] Run-SSLCipherValidation" -ForegroundColor Cyan }
            "6" { Write-Host "[Placeholder] Run-WindowsPatchDetails" -ForegroundColor Cyan }
            "7" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }

        Write-Host "`nPress Enter to return to validation menu..."
        Read-Host | Out-Null
    } while ($true)
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" -ForegroundColor Cyan }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" -ForegroundColor Cyan }
        "4" { exit }
        default { Write-Host "Invalid option." -ForegroundColor Red }
    }
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
        if ($choice -ne "4") {
            Write-Host "`nPress Enter to return to menu..."
            Read-Host | Out-Null
            Clear-Host
        }
    } while ($choice -ne "4")
}

Start-Tool
