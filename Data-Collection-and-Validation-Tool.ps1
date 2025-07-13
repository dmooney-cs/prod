
# Data-Collection-and-Validation-Tool.ps1
# ConnectSecure - System Collection and Validation Launcher (Full with wildcard + safe exit)

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
    Write-Host "`n--- Application Validation ---" -ForegroundColor Cyan
    Write-Host "1. Scan all installed applications"
    Write-Host "2. Scan using a wildcard search term"
    Write-Host "3. Back to Validation Menu"
    $appChoice = Read-Host "Select an option"
    switch ($appChoice) {
        "1" {
            Write-Host "`n[Simulated] Scanning all applications..." -ForegroundColor Green
        }
        "2" {
            $term = Read-Host "Enter keyword or wildcard to search for (e.g. *Office*)"
            Write-Host "`n[Simulated] Searching for: $term" -ForegroundColor Green
        }
        "3" { return }
        default { Write-Host "Invalid option. Returning." -ForegroundColor Red }
    }
    Pause
}

function Run-DriverValidation {
    Write-Host "`n--- Driver Validation ---" -ForegroundColor Cyan
    Write-Host "1. Scan all installed drivers"
    Write-Host "2. Scan using a wildcard search term"
    Write-Host "3. Back to Validation Menu"
    $drvChoice = Read-Host "Select an option"
    switch ($drvChoice) {
        "1" {
            Write-Host "`n[Simulated] Scanning all drivers..." -ForegroundColor Green
        }
        "2" {
            $term = Read-Host "Enter keyword or wildcard to search for (e.g. *NVIDIA*)"
            Write-Host "`n[Simulated] Searching for: $term" -ForegroundColor Green
        }
        "3" { return }
        default { Write-Host "Invalid option. Returning." -ForegroundColor Red }
    }
    Pause
}

function Run-NetworkValidation {
    Write-Host "`n[Simulated] Running Network Validation..." -ForegroundColor Green
    Pause
}

function Run-WindowsUpdateValidation {
    Write-Host "`n[Simulated] Running Windows Update Validation..." -ForegroundColor Green
    Pause
}

function Run-ValidationScripts {
    do {
        Write-Host "`n---- Validation Scripts Menu ----" -ForegroundColor Cyan
        Write-Host "1. Application Validation"
        Write-Host "2. Driver Validation"
        Write-Host "3. Network Validation"
        Write-Host "4. Windows Update Validation"
        Write-Host "5. Back to Main Menu"
        Write-Host "----------------------------------"
        $valChoice = Read-Host "Select an option"

        switch ($valChoice) {
            "1" { Run-ApplicationValidation }
            "2" { Run-DriverValidation }
            "3" { Run-NetworkValidation }
            "4" { Run-WindowsUpdateValidation }
            "5" { return }
            default { Write-Host "Invalid choice. Try again." -ForegroundColor Red }
        }

    } while ($true)
}

function Run-AgentMaintenance {
    Write-Host "`n[Simulated] Agent Maintenance Module Loaded..." -ForegroundColor Cyan
    Pause
}

function Run-ProbeTroubleshooting {
    Write-Host "`n[Simulated] Probe Troubleshooting Module Loaded..." -ForegroundColor Cyan
    Pause
}

function Run-ZipAndEmail {
    Write-Host "`n[Simulated] Zipping Export Folder..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 500
    Write-Host "[Simulated] Launching default email client..." -ForegroundColor Cyan
    Pause
}

function Purge-ScriptData {
    $tempPath = "C:\Script-Temp"
    if (Test-Path $tempPath) {
        $files = Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        $count = 0
        foreach ($item in $files) {
            try {
                Remove-Item -Path $item.FullName -Force -Recurse -ErrorAction Stop
                Write-Host "Deleted: $($item.FullName)" -ForegroundColor DarkGray
                $count++
            } catch {
                Write-Host "Failed to delete: $($item.FullName)" -ForegroundColor Red
            }
        }
        Write-Host "`nTotal items deleted: $count" -ForegroundColor Cyan
    } else {
        Write-Host "No temp folder found at $tempPath." -ForegroundColor Yellow
    }
    Write-Host "`nPress Enter to exit..."
    Read-Host | Out-Null
    exit
}

function Run-SelectedOption {
    param($choice)

    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Run-AgentMaintenance }
        "3" { Run-ProbeTroubleshooting }
        "4" { Run-ZipAndEmail }
        "Q" { Purge-ScriptData }
        default {
            Write-Host "`nInvalid option. Please select again." -ForegroundColor Red
        }
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
        }
        Clear-Host
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
