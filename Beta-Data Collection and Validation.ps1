# Data Collection and Validation Tool - Base Script (for prod and beta)

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

function Run-OfficeValidation {
    Write-Host "Running Microsoft Office Validation..." -ForegroundColor Cyan
    # Placeholder for Office validation code
}

function Run-DriverValidation {
    Write-Host "Running Driver Validation..." -ForegroundColor Cyan
    # Placeholder for Driver validation code
}

function Run-RoamingProfileValidation {
    Write-Host "Running Roaming Profile Applications Validation..." -ForegroundColor Cyan
    # Placeholder for Roaming profile validation code
}

function Run-BrowserExtensionDetails {
    Write-Host "Running Browser Extension Details..." -ForegroundColor Cyan
    # Placeholder for Browser Extension details code
}

function Run-OSQueryBrowserExtensions {
    # Timestamp and hostname
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    $outputCsv = "$outputDir\OSquery-browserext-$timestamp-$hostname.csv"
    $outputJson = "$outputDir\RawBrowserExt-$timestamp-$hostname.json"

    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory | Out-Null
    }

    $osqueryPaths = @(
        "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe",
        "C:\Windows\CyberCNSAgent\osqueryi.exe"
    )

    $osquery = $osqueryPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $osquery) {
        Write-Host "❌ osqueryi.exe not found in expected locations." -ForegroundColor Red
        Read-Host -Prompt "Press any key to exit"
        return
    }

    Push-Location (Split-Path $osquery)

    $sqlQuery = @"
SELECT name, browser_type, version, path, sha1(name || path) AS unique_id 
FROM chrome_extensions 
WHERE chrome_extensions.uid IN (SELECT uid FROM users) 
GROUP BY unique_id;
"@

    $json = & .\osqueryi.exe --json "$sqlQuery" 2>$null
    Pop-Location

    Set-Content -Path $outputJson -Value $json

    if (-not $json) {
        Write-Host "⚠️ No browser extension data returned by osquery." -ForegroundColor Yellow
        Write-Host "`nRaw JSON saved for review:" -ForegroundColor DarkGray
        Write-Host "$outputJson" -ForegroundColor Cyan
        Read-Host -Prompt "Press any key to continue"
        return
    }

    try {
        $parsed = $json | ConvertFrom-Json

        if ($parsed.Count -eq 0) {
            Write-Host "⚠️ No browser extensions found." -ForegroundColor Yellow
        } else {
            Write-Host "`n✅ Extensions found: $($parsed.Count)" -ForegroundColor Green
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            foreach ($ext in $parsed) {
                Write-Host "• $($ext.name) ($($ext.browser_type)) — $($ext.version)" -ForegroundColor White
            }
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            $parsed | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
        }
    } catch {
        Write-Host "❌ Failed to parse or export data." -ForegroundColor Red
        Read-Host -Prompt "Press any key to continue"
        return
    }

    Write-Host "`n📁 Output Files:" -ForegroundColor Yellow
    Write-Host "• CSV:  $outputCsv" -ForegroundColor Cyan
    Write-Host "• JSON: $outputJson" -ForegroundColor Cyan
    Read-Host -Prompt "`nPress any key to exit"
}

function Run-SSLCipherValidation {
    Write-Host "Running SSL Cipher Validation..." -ForegroundColor Cyan
    # Placeholder for SSL cipher validation code
}

function Run-WindowsPatchDetails {
    Write-Host "Running Windows Patch Details..." -ForegroundColor Cyan
    # Placeholder for Windows patch details code
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Write-Host "[Placeholder] Zip and Email Results" }
        "Q" { exit }
        default { Write-Host "Invalid option." -ForegroundColor Red }
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
# Data Collection and Validation Tool - Base Script (for prod and beta)

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

function Run-OfficeValidation {
    Write-Host "Running Microsoft Office Validation..." -ForegroundColor Cyan
    # Placeholder for Office validation code
}

function Run-DriverValidation {
    Write-Host "Running Driver Validation..." -ForegroundColor Cyan
    # Placeholder for Driver validation code
}

function Run-RoamingProfileValidation {
    Write-Host "Running Roaming Profile Applications Validation..." -ForegroundColor Cyan
    # Placeholder for Roaming profile validation code
}

function Run-BrowserExtensionDetails {
    Write-Host "Running Browser Extension Details..." -ForegroundColor Cyan
    # Placeholder for Browser Extension details code
}

function Run-OSQueryBrowserExtensions {
    # Timestamp and hostname
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $outputDir = "C:\Script-Export"
    $outputCsv = "$outputDir\OSquery-browserext-$timestamp-$hostname.csv"
    $outputJson = "$outputDir\RawBrowserExt-$timestamp-$hostname.json"

    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory | Out-Null
    }

    $osqueryPaths = @(
        "C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe",
        "C:\Windows\CyberCNSAgent\osqueryi.exe"
    )

    $osquery = $osqueryPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $osquery) {
        Write-Host "❌ osqueryi.exe not found in expected locations." -ForegroundColor Red
        Read-Host -Prompt "Press any key to exit"
        return
    }

    Push-Location (Split-Path $osquery)

    $sqlQuery = @"
SELECT name, browser_type, version, path, sha1(name || path) AS unique_id 
FROM chrome_extensions 
WHERE chrome_extensions.uid IN (SELECT uid FROM users) 
GROUP BY unique_id;
"@

    $json = & .\osqueryi.exe --json "$sqlQuery" 2>$null
    Pop-Location

    Set-Content -Path $outputJson -Value $json

    if (-not $json) {
        Write-Host "⚠️ No browser extension data returned by osquery." -ForegroundColor Yellow
        Write-Host "`nRaw JSON saved for review:" -ForegroundColor DarkGray
        Write-Host "$outputJson" -ForegroundColor Cyan
        Read-Host -Prompt "Press any key to continue"
        return
    }

    try {
        $parsed = $json | ConvertFrom-Json

        if ($parsed.Count -eq 0) {
            Write-Host "⚠️ No browser extensions found." -ForegroundColor Yellow
        } else {
            Write-Host "`n✅ Extensions found: $($parsed.Count)" -ForegroundColor Green
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            foreach ($ext in $parsed) {
                Write-Host "• $($ext.name) ($($ext.browser_type)) — $($ext.version)" -ForegroundColor White
            }
            Write-Host "--------------------------------------------------" -ForegroundColor DarkGray
            $parsed | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
        }
    } catch {
        Write-Host "❌ Failed to parse or export data." -ForegroundColor Red
        Read-Host -Prompt "Press any key to continue"
        return
    }

    Write-Host "`n📁 Output Files:" -ForegroundColor Yellow
    Write-Host "• CSV:  $outputCsv" -ForegroundColor Cyan
    Write-Host "• JSON: $outputJson" -ForegroundColor Cyan
    Read-Host -Prompt "`nPress any key to exit"
}

function Run-SSLCipherValidation {
    Write-Host "Running SSL Cipher Validation..." -ForegroundColor Cyan
    # Placeholder for SSL cipher validation code
}

function Run-WindowsPatchDetails {
    Write-Host "Running Windows Patch Details..." -ForegroundColor Cyan
    # Placeholder for Windows patch details code
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Write-Host "[Placeholder] Zip and Email Results" }
        "Q" { exit }
        default { Write-Host "Invalid option." -ForegroundColor Red }
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