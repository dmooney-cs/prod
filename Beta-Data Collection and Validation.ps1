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

function Run-BrowserExtensionDetails {
    function Get-ChromeEdgeExtensions {
        param (
            [string]$BrowserName,
            [string]$BasePath
        )
        if (-Not (Test-Path $BasePath)) { return }
        Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $extensionId = $_.Name
            Get-ChildItem -Path $_.FullName -Directory | Sort-Object Name -Descending | Select-Object -First 1 | ForEach-Object {
                $manifestPath = Join-Path $_.FullName 'manifest.json'
                if (Test-Path $manifestPath) {
                    try {
                        $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                        [PSCustomObject]@{
                            Browser         = $BrowserName
                            ExtensionID     = $extensionId
                            Name            = $manifest.name
                            Version         = $manifest.version
                            Description     = $manifest.description
                            Path            = $_.FullName
                            InstallLocation = $BasePath
                        }
                    } catch {}
                }
            }
        }
    }

    function Get-FirefoxExtensions {
        param ([string]$UserProfile)
        $firefoxProfilesIni = Join-Path $UserProfile 'AppData\Roaming\Mozilla\Firefox\profiles.ini'
        if (-Not (Test-Path $firefoxProfilesIni)) { return }

        $profileDirs = Select-String -Path $firefoxProfilesIni -Pattern '^Path=' | ForEach-Object {
            $_.Line -replace 'Path=', ''
        }

        foreach ($profileDir in $profileDirs) {
            $extensionsPath = Join-Path $UserProfile "AppData\Roaming\Mozilla\Firefox\$profileDir\extensions"
            if (Test-Path $extensionsPath) {
                Get-ChildItem -Path $extensionsPath -File | ForEach-Object {
                    [PSCustomObject]@{
                        Browser         = 'Firefox'
                        ExtensionID     = $_.Name
                        Name            = ''
                        Version         = ''
                        Description     = ''
                        Path            = $_.FullName
                        InstallLocation = $extensionsPath
                    }
                }
            }
        }
    }

    $AllResults = @()
    $Users = Get-ChildItem 'C:\Users' -Directory | Where-Object { $_.Name -notin @('Default', 'Public', 'All Users') }

    foreach ($user in $Users) {
        $profilePath = $user.FullName

        $chromePath = Join-Path $profilePath 'AppData\Local\Google\Chrome\User Data\Default\Extensions'
        $AllResults += Get-ChromeEdgeExtensions -BrowserName 'Chrome' -BasePath $chromePath

        $edgePath = Join-Path $profilePath 'AppData\Local\Microsoft\Edge\User Data\Default\Extensions'
        $AllResults += Get-ChromeEdgeExtensions -BrowserName 'Edge' -BasePath $edgePath

        $AllResults += Get-FirefoxExtensions -UserProfile $profilePath
    }

    $SortedResults = $AllResults | Sort-Object Browser

    $SortedResults | Format-Table -Property Browser, ExtensionID, Name, Version, Description, InstallLocation, Path -AutoSize

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $hostname = $env:COMPUTERNAME
    $exportPath = "C:\Script-Export"

    if (-Not (Test-Path $exportPath)) {
        New-Item -Path $exportPath -ItemType Directory | Out-Null
    }

    $csvPath = Join-Path $exportPath "BrowserExtensions_${timestamp}_${hostname}.csv"
    $SortedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport saved to: $csvPath" -ForegroundColor Green
    Read-Host -Prompt "`nPress ENTER to return to menu"
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Write-Host "[Placeholder] Agent Maintenance" }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Write-Host "[Placeholder] Zip and Email Results" }
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
