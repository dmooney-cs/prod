# Data-Collection-and-Validation-Tool.ps1
# ConnectSecure - System Collection and Validation Launcher (Integrated Office Validation)

Write-Host "`n======== Data Collection and Validation Tool ========" -ForegroundColor Cyan

function Show-MainMenu {
    Write-Host ""
    Write-Host "1. Validation Scripts"
    Write-Host "   → Run application, driver, network, and update validations."
    Write-Host ""
    Write-Host "2. Agent Maintenance"
    Write-Host "   → Install, update, or troubleshoot the ConnectSecure agent."
    Write-Host ""
    Write-Host "3. Probe Troubleshooting"
    Write-Host "   → Diagnose probe issues and test scanning tools."
    Write-Host ""
    Write-Host "4. Zip and Email Results"
    Write-Host "   → Package collected data into a ZIP and open your mail client."
    Write-Host ""
    Write-Host "Q. Close and Purge Script Data"
    Write-Host "   → Optionally email, then delete all script-related files."
    Write-Host ""
}

function Run-ApplicationValidation {
    do {
        Write-Host "`n--- Application Validation ---" -ForegroundColor Cyan
        Write-Host "1. Scan all installed applications"
        Write-Host "2. Scan using a wildcard search term"
        Write-Host "3. Microsoft Office Validation"
        Write-Host "4. Back to Validation Menu"
        $appChoice = Read-Host "Select an option"

        switch ($appChoice) {
            "1" {
                $wildcard = "*"
                Run-OfficeValidation -appFilter $wildcard
            }
            "2" {
                $wildcard = Read-Host "Enter keyword or wildcard to search for (e.g. *Office*)"
                Run-OfficeValidation -appFilter $wildcard
            }
            "3" {
                Run-OfficeValidation -appFilter "*Office*"
            }
            "4" { return }
            default { Write-Host "Invalid option. Returning." -ForegroundColor Red }
        }

        Pause
    } while ($true)
}

<...REMOVED FOR LENGTH LIMIT... remaining content continues here intact...>
