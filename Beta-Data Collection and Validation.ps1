
<#
Data Collection and Validation Tool - Master Script
All Logic Inlined: Validation, Agent Maintenance, Troubleshooting, and Export
#>

# Create export directory
$ExportDir = "C:\Script-Export"
if (-not (Test-Path $ExportDir)) {
    New-Item -Path $ExportDir -ItemType Directory | Out-Null
}

function Pause-Script {
    Write-Host "`nPress any key to continue..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

function Show-MainMenu {
    Clear-Host
    Write-Host "======== Data Collection and Validation Tool ========" -ForegroundColor Cyan
    Write-Host "1. Validation Scripts"
    Write-Host "2. Agent Maintenance"
    Write-Host "3. Probe Troubleshooting"
    Write-Host "4. Zip and Email Results"
    Write-Host "Q. Close and Purge Script Data"
    $selection = Read-Host "Select an option"
    switch ($selection) {
        '1' { Show-ValidationMenu }
        '2' { Show-AgentMaintenanceMenu }
        '3' { Show-ProbeTroubleshootingMenu }
        '4' { Run-ZipAndEmailResults }
        'Q' { Cleanup-And-Exit }
        default { Show-MainMenu }
    }
}

function Show-ValidationMenu {
    Clear-Host
    Write-Host "======== Validation Scripts Menu ========" -ForegroundColor Cyan
    Write-Host "1. Microsoft Office Validation"
    Write-Host "2. Driver Validation"
    Write-Host "3. Roaming Profile Applications"
    Write-Host "4. Browser Extension Details"
    Write-Host "5. OSQuery Browser Extensions"
    Write-Host "6. SSL Cipher Validation"
    Write-Host "7. Windows Patch Details"
    Write-Host "8. Active Directory Collection"
    Write-Host "9. Back to Main Menu"
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Run-OfficeValidation }
        '2' { Run-DriverValidation }
        '3' { Run-RoamingProfileApps }
        '4' { Run-BrowserExtensionDetails }
        '5' { Run-OSQueryBrowserExtensions }
        '6' { Run-SSLCipherValidation }
        '7' { Run-WindowsPatchValidation }
        '8' { Run-ActiveDirectoryValidation }
        '9' { Show-MainMenu }
        default { Show-ValidationMenu }
    }
}

function Run-OfficeValidation {
    Write-Host "Scanning for installed applications..." -ForegroundColor Cyan
    $appFilter = Read-Host "Enter a keyword to filter applications (or press Enter to list all)"

    $results = @()

    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($regPath in $registryPaths) {
        Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.DisplayName) {
                if ($appFilter -eq "" -or $_.DisplayName -like "*$appFilter*") {
                    $category = "Application"
                    $registryPath = $_.PSPath.Replace("Microsoft.PowerShell.Core\Registry::", "")
                    $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }

                    $results += [PSCustomObject]@{
                        Name             = $_.DisplayName
                        Version          = $_.DisplayVersion
                        Publisher        = $_.Publisher
                        InstallLocation  = $installLocation
                        InstallDate      = $_.InstallDate
                        InstalledBy      = $_.InstallSource
                        Source           = "Registry"
                        Category         = $category
                        RegistryPath     = $registryPath
                        ProfileType      = "N/A"
                    }
                }
            }
        }
    }

    Get-AppxPackage | ForEach-Object {
        if ($appFilter -eq "" -or $_.Name -like "*$appFilter*" -or $_.PackageFullName -like "*$appFilter*") {
            $category = "Microsoft Store App"
            $installLocation = if ($_.InstallLocation) { $_.InstallLocation } else { "N/A" }

            $results += [PSCustomObject]@{
                Name             = $_.Name
                Version          = $_.Version
                Publisher        = $_.Publisher
                InstallLocation  = $installLocation
                InstallDate      = "N/A"
                InstalledBy      = "N/A"
                Source           = "Microsoft Store"
                Category         = $category
                RegistryPath     = "N/A"
                ProfileType      = "N/A"
            }
        }
    }

    $teamsPaths = Get-ChildItem "C:\Users\*\AppData\Local\Microsoft\Teams\Teams.exe" -Recurse -ErrorAction SilentlyContinue
    foreach ($teamsPath in $teamsPaths) {
        $version = (Get-Item $teamsPath.FullName).VersionInfo.FileVersion
        $installLocation = $teamsPath.DirectoryName

        $results += [PSCustomObject]@{
            Name             = 'Microsoft Teams'
            Version          = $version
            Publisher        = 'Microsoft'
            InstallLocation  = $installLocation
            InstallDate      = "N/A"
            InstalledBy      = "N/A"
            Source           = "Teams (Local)"
            Category         = "Microsoft Store App"
            RegistryPath     = "N/A"
            ProfileType      = "N/A"
        }
    }

    if ($results.Count -eq 0) {
        Write-Host "No matching applications found." -ForegroundColor Yellow
    } else {
        Write-Host "`nDetected Applications:" -ForegroundColor Green
        $results | Sort-Object Name | Format-Table Name, Version, Publisher, InstallLocation, InstallDate, InstalledBy, Category, Source -AutoSize

        $dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $hostname = $env:COMPUTERNAME
        $exportPath = "C:\Script-Export"

        if (-not (Test-Path -Path $exportPath)) {
            New-Item -Path $exportPath -ItemType Directory
            Write-Host "Created folder: $exportPath" -ForegroundColor Cyan
        }

        $csvPath = "$exportPath\All Applications Detected - $dateTime - $hostname.csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nResults exported to: $csvPath" -ForegroundColor Green
    }
    Pause-Script
}

# Import Agent Maintenance Submenu
. { <INSERT AGENT MAINTENANCE MENU V1 SCRIPT HERE> }

# Import Probe Troubleshooting Submenu
. { <INSERT PROBE TROUBLESHOOTING MENU V1 SCRIPT HERE> }

# Import Zip + Email Logic
. { <INSERT ZIP AND EMAIL RESULTS V1 SCRIPT HERE> }

# Purge Export and Temp Data
function Cleanup-And-Exit {
    Write-Host "`nCleaning up all Script Data..." -ForegroundColor Yellow
    $pathsToDelete = @("C:\Script-Export", "C:\Script-Temp")
    foreach ($path in $pathsToDelete) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Host "Script data purged successfully. Exiting..." -ForegroundColor Green
    Pause-Script
    exit
}

# Launch Menu
Show-MainMenu
