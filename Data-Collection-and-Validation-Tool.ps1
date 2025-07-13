
# Data-Collection-and-Validation-Tool.ps1 (Standalone)
# Fully self-contained script including all validation options

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

function Run-ValidationScripts {
    Write-Host "`n[Simulated] Running Application Validation..." -ForegroundColor Green
    Start-Sleep -Milliseconds 500
    Write-Host "[Simulated] Running Driver Validation..." -ForegroundColor Green
    Start-Sleep -Milliseconds 500
    Write-Host "[Simulated] Running Network Validation..." -ForegroundColor Green
    Start-Sleep -Milliseconds 500
    Write-Host "[Simulated] Running Windows Update Validation..." -ForegroundColor Green
    Start-Sleep -Milliseconds 500
    Write-Host "`nAll validations completed." -ForegroundColor Cyan
}

function Run-AgentMaintenance {
    Write-Host "`n[Simulated] Agent Maintenance Module Loaded..." -ForegroundColor Cyan
}

function Run-ProbeTroubleshooting {
    Write-Host "`n[Simulated] Probe Troubleshooting Module Loaded..." -ForegroundColor Cyan
}

function Run-ZipAndEmail {
    Write-Host "`n[Simulated] Zipping Export Folder..." -ForegroundColor Cyan
    Start-Sleep -Milliseconds 500
    Write-Host "[Simulated] Launching default email client..." -ForegroundColor Cyan
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
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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
            [void][System.Console]::ReadLine()
        }
        Clear-Host
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
