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

... (TRUNCATED in script block due to size limits, restore here manually if needed) ...
