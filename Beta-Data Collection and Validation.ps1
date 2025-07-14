# Data-Collection-and-Validation-Tool.ps1

Clear-Host
Write-Host "=== Data Collection and Validation Tool ===" -ForegroundColor Cyan

function Show-MainMenu {
    Clear-Host
    Write-Host "=== Data Collection and Validation Tool ===" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
    Write-Host ""
}

function Show-ValidationMenu {
    Clear-Host
    Write-Host "=== Validation Scripts ===" -ForegroundColor Cyan
    Write-Host "1. Microsoft Office Validation"
    Write-Host "2. Driver Validation"
    Write-Host "3. Roaming Profile Applications"
    Write-Host "4. Browser Extension Details"
    Write-Host "5. OSQuery Browser Extensions"
    Write-Host "6. SSL Cipher Validation"
    Write-Host "7. Windows Patch Details"
    Write-Host "8. Active Directory Collection"
    Write-Host "9. Back to Main Menu"
    Write-Host ""
}

function Run-BrowserExtensionAudit {
    function Get-ChromeEdgeExtensions {
        param (
            [string]$BrowserName,
            [string]$BasePath
        )

        if (-Not (Test-Path $BasePath)) {
            return
        }

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
        param (
            [string]$UserProfile
        )

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
        Write-Host "Creating folder: $exportPath"
        New-Item -Path $exportPath -ItemType Directory
    }

    $csvPath = Join-Path $exportPath "BrowserExtensions_$timestamp`_$hostname.csv"
    $SortedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport saved to: $csvPath" -ForegroundColor Green
    Pause
}

function Run-AllValidation {
    Run-OfficeValidation
    Run-DriverValidation
    Run-RoamingProfileAppScan
    Run-BrowserExtensionAudit
    Run-OSQueryBrowserExt
    Run-SSLCipherScan
    Run-WindowsPatchDetails
    Run-ADCollection
}

function Run-OfficeValidation { # Stub - Replace with real logic }
function Run-DriverValidation { # Stub - Replace with real logic }
function Run-RoamingProfileAppScan { # Stub - Replace with real logic }
function Run-OSQueryBrowserExt { # Stub - Replace with real logic }
function Run-SSLCipherScan { # Stub - Replace with real logic }
function Run-WindowsPatchDetails { # Stub - Replace with real logic }
function Run-ADCollection { # Stub - Replace with real logic }
function Run-ZipAndEmail { # Stub - Replace with real logic }
function Run-AgentMaintenance { # Stub - Replace with real logic }
function Run-ProbeTroubleshooting { # Stub - Replace with real logic }

# Main Loop
do {
    Show-MainMenu
    $mainChoice = Read-Host "Select an option"

    switch ($mainChoice) {
        '1' {
            do {
                Show-ValidationMenu
                $validationChoice = Read-Host "Choose a validation option"

                switch ($validationChoice) {
                    '1' { Run-OfficeValidation }
                    '2' { Run-DriverValidation }
                    '3' { Run-RoamingProfileAppScan }
                    '4' { Run-BrowserExtensionAudit }
                    '5' { Run-OSQueryBrowserExt }
                    '6' { Run-SSLCipherScan }
                    '7' { Run-WindowsPatchDetails }
                    '8' { Run-ADCollection }
                    '9' { break }
                    default { Write-Host "Invalid option. Try again." -ForegroundColor Red; Pause }
                }
            } while ($true)
        }
        '2' { Run-AgentMaintenance }
        '3' { Run-ProbeTroubleshooting }
        '4' { Run-ZipAndEmail }
        'Q' { break }
        default { Write-Host "Invalid option. Try again." -ForegroundColor Red; Pause }
    }
} while ($true)

Write-Host "Cleaning up and exiting..." -ForegroundColor Yellow
