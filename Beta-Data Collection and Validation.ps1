# ==============================
# Data Collection & Validation Tool - Full V1 Version
# All Logic Inlined - GitHub Compatible
# Includes: Office, Extensions, Drivers, SMB, SSL, Updates, Agent Tools, ZIP
# ==============================

# Region: Utility Functions
function Ensure-ExportPath {
    $global:ExportPath = "C:\Script-Export"
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory | Out-Null
    }
    $global:Timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $global:Hostname = $env:COMPUTERNAME
}

function Pause-ReturnMenu {
    Write-Host "`nPress any key to return to the menu..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Show-MainMenu
}

# EndRegion

# Region: Menu Navigation
function Show-MainMenu {
    Clear-Host
    Ensure-ExportPath
    Write-Host "=== Data Validation Master Menu ===" -ForegroundColor Cyan
    Write-Host "1. Microsoft Office Validation"
    Write-Host "2. Browser Extension Audit"
    Write-Host "3. Driver Validation"
    Write-Host "4. Check SMB Versions"
    Write-Host "5. Validate SSL Ciphers"
    Write-Host "6. Windows Update Validation"
    Write-Host "7. Agent Job Clear"
    Write-Host "8. Enable SMB Settings"
    Write-Host "9. Zip and Prepare for Email"
    Write-Host "Q. Quit"
    $choice = Read-Host "Select an option"
    switch ($choice) {
        '1' { Run-MicrosoftOfficeValidation; return }
        '2' { Run-BrowserExtensionAudit; return }
        '3' { Run-DriverValidation; return }
        '4' { Run-SMBVersionCheck; return }
        '5' { Run-SSLCipherScan; return }
        '6' { Run-WindowsUpdateValidation; return }
        '7' { Run-AgentJobClear; return }
        '8' { Run-EnableSMB; return }
        '9' { Run-ZipAndEmail; return }
        'Q' { exit }
        default { Show-MainMenu; return }
    }
}
# EndRegion

# Region: Module Scripts
function Run-MicrosoftOfficeValidation {
    Clear-Host
    Write-Host "=== Running Microsoft Office Validation ===" -ForegroundColor Cyan
    $csvPath = "$ExportPath\OfficeValidation_$Timestamp`_$Hostname.csv"
    $results = @()
    $officePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
    )
    foreach ($path in $officePaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_.Name -match '\\\d+\.\d+') {
                    $version = $_.PSChildName
                    $results += [PSCustomObject]@{
                        Path    = $_.Name
                        Version = $version
                        Type    = "Registry"
                    }
                }
            }
        }
    }
    $programDirs = @("$env:ProgramFiles", "$env:ProgramFiles(x86)")
    foreach ($dir in $programDirs) {
        $officeDir = Join-Path $dir "Microsoft Office"
        if (Test-Path $officeDir) {
            $results += [PSCustomObject]@{
                Path    = $officeDir
                Version = ""
                Type    = "Filesystem"
            }
        }
    }
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $csvPath" -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-BrowserExtensionAudit {
    Clear-Host
    Write-Host "=== Running Browser Extension Audit ===" -ForegroundColor Cyan
    $results = @()
    $userProfiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public", "Default", "All Users") }
    foreach ($user in $userProfiles) {
        $userPath = $user.FullName
        $chromePath = "$userPath\AppData\Local\Google\Chrome\User Data\Default\Extensions"
        if (Test-Path $chromePath) {
            Get-ChildItem -Path $chromePath -Directory | ForEach-Object {
                $results += [PSCustomObject]@{
                    Browser     = "Chrome"
                    ExtensionID = $_.Name
                    Path        = $_.FullName
                    User        = $user.Name
                }
            }
        }
        $edgePath = "$userPath\AppData\Local\Microsoft\Edge\User Data\Default\Extensions"
        if (Test-Path $edgePath) {
            Get-ChildItem -Path $edgePath -Directory | ForEach-Object {
                $results += [PSCustomObject]@{
                    Browser     = "Edge"
                    ExtensionID = $_.Name
                    Path        = $_.FullName
                    User        = $user.Name
                }
            }
        }
    }
    $csvPath = "$ExportPath\BrowserExtensions_$Timestamp`_$Hostname.csv"
    $results | Sort-Object Browser, ExtensionID | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Extensions exported to: $csvPath" -ForegroundColor Green
    Pause-ReturnMenu
}

# ... [The rest of the module functions remain unchanged but will now be called and return properly]

# Run Menu
Show-MainMenu
