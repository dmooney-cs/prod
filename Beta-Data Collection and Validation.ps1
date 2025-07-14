function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Microsoft Office Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Roaming Profile Applications"
        Write-Host "4. Browser Extension Details"
        Write-Host "5. OSQuery Browser Extensions"
        Write-Host "6. SSL Cipher Validation"
        Write-Host "7. Windows Patch Details"
        Write-Host "8. Back to Main Menu"
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

function Run-AgentMaintenance {
    do {
        Write-Host "`n---- Agent Maintenance Menu ----" -ForegroundColor Cyan
        Write-Host "1. Agent - Clear Jobs"
        Write-Host "2. Agent - Install Utility"
        Write-Host "3. Agent - Check SMB"
        Write-Host "4. Agent - Set SMB"
        Write-Host "5. Back to Main Menu"
        $agentChoice = Read-Host "Select an option"
        switch ($agentChoice) {
            "1" { Run-AgentClearJobs }
            "2" { Run-AgentInstallUtility }
            "3" { Run-AgentCheckSMB }
            "4" { Run-AgentSetSMB }
            "5" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-AgentClearJobs {
    Write-Host "[AgentClearJobs executed]" -ForegroundColor Green
}

function Run-AgentInstallUtility {
    Write-Host "[AgentInstallUtility executed]" -ForegroundColor Green
}

function Run-AgentCheckSMB {
    Write-Host "[AgentCheckSMB executed]" -ForegroundColor Green
}

function Run-AgentSetSMB {
    Write-Host "[AgentSetSMB executed]" -ForegroundColor Green
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Run-AgentMaintenance }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Run-ZipAndEmailResults }
        "Q" { exit }
    }
}

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
