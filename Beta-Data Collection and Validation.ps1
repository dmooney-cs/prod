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
        '1' { Run-MicrosoftOfficeValidation }
        '2' { Run-BrowserExtensionAudit }
        '3' { Run-DriverValidation }
        '4' { Run-SMBVersionCheck }
        '5' { Run-SSLCipherScan }
        '6' { Run-WindowsUpdateValidation }
        '7' { Run-AgentJobClear }
        '8' { Run-EnableSMB }
        '9' { Run-ZipAndEmail }
        'Q' { exit }
        default { Show-MainMenu }
    }
}
# EndRegion

# Region: Module Scripts
function Run-MicrosoftOfficeValidation {
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

function Run-DriverValidation {
    $results = Get-WmiObject Win32_PnPSignedDriver | Select-Object DeviceName, DriverVersion, Manufacturer, DriverProviderName, DriverDate
    $csvPath = "$ExportPath\DriverValidation_$Timestamp`_$Hostname.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Drivers exported to: $csvPath" -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-SMBVersionCheck {
    $results = @()
    $regKeys = @(
        @{ Name = "SMBv1"; Path = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"; Key = "SMB1" },
        @{ Name = "SMBv2"; Path = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"; Key = "SMB2" }
    )
    foreach ($reg in $regKeys) {
        $value = Get-ItemProperty -Path $reg.Path -Name $reg.Key -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $reg.Key -ErrorAction SilentlyContinue
        $results += [PSCustomObject]@{
            Protocol     = $reg.Name
            Enabled      = if ($value -eq 1) { "Yes" } else { "No" }
            RegistryPath = $reg.Path
            Key          = $reg.Key
            DisableKey   = "Set-ItemProperty -Path '$($reg.Path)' -Name '$($reg.Key)' -Value 0"
        }
    }
    $csvPath = "$ExportPath\SMB_Version_Report_$Timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "SMB status exported to: $csvPath" -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-SSLCipherScan {
    Write-Host "Launching SSL cipher scan using Nmap..." -ForegroundColor Cyan
    $nmapPath = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmapPath)) {
        Write-Host "Nmap not found: $nmapPath" -ForegroundColor Red
        Pause-ReturnMenu
        return
    }
    $target = Read-Host "Enter the target IP address for scan"
    $outputPath = "$ExportPath\TLS443Scan-$target-$Timestamp-$Hostname.csv"
    Push-Location (Split-Path $nmapPath)
    .\nmap.exe --script ssl-enum-ciphers -p 443 $target | Out-File -FilePath $outputPath
    Pop-Location
    Write-Host "Scan results saved to: $outputPath" -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-WindowsUpdateValidation {
    $outputCSV = "$ExportPath\WindowsUpdates_$Timestamp`_$Hostname.csv"
    $results = Get-HotFix | Select-Object HotFixID, Description, InstalledOn, InstalledBy
    $results | Export-Csv -Path $outputCSV -NoTypeInformation -Encoding UTF8
    Write-Host "Windows Updates exported to: $outputCSV" -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-AgentJobClear {
    Write-Host "Running Agent Check Job Clear..." -ForegroundColor Cyan
    $agentDir = "C:\Program Files (x86)\CyberCNSAgent"
    $exePath = "$agentDir\agentcheck.exe"
    $tempDir = "C:\Script-Temp"
    if (-not (Test-Path $tempDir)) { New-Item $tempDir -ItemType Directory | Out-Null }
    if (-not (Test-Path $exePath)) {
        Invoke-WebRequest -Uri "https://agentv3.myconnectsecure.com/agentcheck.exe" -OutFile "$tempDir\agentcheck.exe"
        Copy-Item "$tempDir\agentcheck.exe" -Destination $agentDir -Force
    }
    Stop-Service cybercnsagent -ErrorAction SilentlyContinue
    Stop-Service cybercnsagentmonitor -ErrorAction SilentlyContinue
    Remove-Item "$agentDir\pendingjobqueue\*" -Recurse -Force -ErrorAction SilentlyContinue
    Push-Location $agentDir
    .\agentcheck.exe
    Pop-Location
    Start-Service cybercnsagent
    Start-Service cybercnsagentmonitor
    Write-Host "Agent check complete." -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-EnableSMB {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" SMB2 -Type DWORD -Value 1 -Force
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (SMB-In)" -Enabled true -Profile Any
    Set-NetFirewallRule -DisplayName "File And Printer Sharing (NB-Session-In)" -Enabled true -Profile Any
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f | Out-Null
    Write-Host "SMB has been enabled and configured." -ForegroundColor Green
    Pause-ReturnMenu
}

function Run-ZipAndEmail {
    $zipPath = "$ExportPath\ScriptResults_$Timestamp.zip"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path "$ExportPath\*" -DestinationPath $zipPath
    Write-Host "ZIP created: $zipPath" -ForegroundColor Yellow
    Start-Process "mailto:support@connectsecure.com?subject=Validation Report $Hostname&body=Please attach the ZIP file: $zipPath"
    Pause-ReturnMenu
}

# EndRegion

# Run Menu
Show-MainMenu
