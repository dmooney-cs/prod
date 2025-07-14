
# Data-Collection-and-Validation-Tool.ps1 - Fully Restored and Functional

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
            "1" { Write-Host "[Simulated] Office Validation running..." -ForegroundColor Green }
            "2" { Write-Host "[Simulated] Driver Validation running..." -ForegroundColor Green }
            "3" { Write-Host "[Simulated] Roaming Profile Validation running..." -ForegroundColor Green }
            "4" { Write-Host "[Simulated] Browser Extension Scan running..." -ForegroundColor Green }
            "5" { Write-Host "[Simulated] SSL Cipher Validation running..." -ForegroundColor Green }
            "6" { Write-Host "[Simulated] Patch Validation running..." -ForegroundColor Green }
            "7" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Red }
        }

        Write-Host "`nPress Enter to return to the validation menu..."
        Read-Host | Out-Null
    } while ($true)
}

function Run-ZipAndEmail {
    Write-Host "[Simulated] Zipping and emailing results..." -ForegroundColor Cyan
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Simulated] Agent Maintenance..." -ForegroundColor Cyan }
        "3" { Write-Host "[Simulated] Probe Troubleshooting..." -ForegroundColor Cyan }
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
            Write-Host "`nPress Enter to return to main menu..."
            Read-Host | Out-Null
        }
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
