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

function Run-AgentMaintenance {
    do {
        Write-Host "`n---- Agent Maintenance Menu ----" -ForegroundColor Cyan
        Write-Host "1. Agent - Clear Jobs"
        Write-Host "2. Agent - Install Utility"
        Write-Host "3. Agent - Check SMB"
        Write-Host "4. Agent - Set SMB"
        Write-Host "5. Back to Main Menu"
        $agentChoice = Read-Host "Select an option"
        switch ($agentChoice) {
            "1" { Run-AgentClearJobs }
            "2" { Run-AgentInstallUtility }
            "3" { Run-AgentCheckSMB }
            "4" { Run-AgentSetSMB }
            "5" { return }
            default { Write-Host "Invalid option." -ForegroundColor Red }
        }
    } while ($true)
}

function Run-AgentClearJobs {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $exportDir = "C:\Script-Export"
    $transcriptFile = "$exportDir\AgentCheck-FullOutput-$timestamp-$hostname.txt"
    $csvLogFile = "$exportDir\AgentCheck-Summary-$timestamp-$hostname.csv"
    $tempDir = "C:\Script-Temp"
    $agentDir = "C:\Program Files (x86)\CyberCNSAgent"
    $pendingJobQueue = Join-Path $agentDir "pendingjobqueue"
    $downloadUrl = "https://agentv3.myconnectsecure.com/agentcheck.exe"
    $downloadedFile = Join-Path $tempDir "agentcheck.exe"
    $agentFileDest = Join-Path $agentDir "agentcheck.exe"

    $log = @()

    foreach ($folder in @($exportDir, $tempDir)) {
        if (-not (Test-Path $folder)) {
            New-Item -Path $folder -ItemType Directory | Out-Null
            $log += [PSCustomObject]@{ Step = "Create Folder"; Detail = $folder; Status = "Created" }
        } else {
            $log += [PSCustomObject]@{ Step = "Check Folder"; Detail = $folder; Status = "Exists" }
        }
    }

    Start-Transcript -Path $transcriptFile -Force

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadedFile -UseBasicParsing
        Copy-Item -Path $downloadedFile -Destination $agentFileDest -Force
        Write-Host "Downloaded and copied agentcheck.exe to agent folder." -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Download & Copy"; Detail = $agentFileDest; Status = "Success" }
    } catch {
        Write-Host "Failed to download or copy agentcheck.exe." -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Download & Copy"; Detail = $_.Exception.Message; Status = "Failed" }
    }

    foreach ($svc in @("cybercnsagent", "cybercnsagentmonitor")) {
        try {
            Stop-Service -Name $svc -Force -ErrorAction Stop
            Write-Host "Stopped service: $svc" -ForegroundColor Green
            $log += [PSCustomObject]@{ Step = "Stop Service"; Detail = $svc; Status = "Stopped" }
        } catch {
            Write-Host "Failed to stop service: $svc" -ForegroundColor Yellow
            $log += [PSCustomObject]@{ Step = "Stop Service"; Detail = $svc; Status = "Not Running or Failed" }
        }
    }

    if (Test-Path $pendingJobQueue) {
        try {
            Remove-Item "$pendingJobQueue\*" -Recurse -Force -ErrorAction Stop
            Write-Host "Cleared pendingjobqueue." -ForegroundColor Green
            $log += [PSCustomObject]@{ Step = "Clear Queue"; Detail = $pendingJobQueue; Status = "Cleared" }
        } catch {
            Write-Host "Failed to clear queue: $_" -ForegroundColor Yellow
            $log += [PSCustomObject]@{ Step = "Clear Queue"; Detail = $_.Exception.Message; Status = "Failed" }
        }
    } else {
        Write-Host "pendingjobqueue directory not found." -ForegroundColor Yellow
        $log += [PSCustomObject]@{ Step = "Clear Queue"; Detail = $pendingJobQueue; Status = "Not Found" }
    }

    try {
        Write-Host "Running agentcheck.exe..."
        & "$agentFileDest"
        Write-Host "agentcheck.exe execution complete." -ForegroundColor Green
        $log += [PSCustomObject]@{ Step = "Run AgentCheck"; Detail = $agentFileDest; Status = "Executed" }
    } catch {
        Write-Host "Failed to run agentcheck.exe." -ForegroundColor Red
        $log += [PSCustomObject]@{ Step = "Run AgentCheck"; Detail = $_.Exception.Message; Status = "Failed" }
    }

    foreach ($svc in @("cybercnsagent", "cybercnsagentmonitor")) {
        try {
            Start-Service -Name $svc -ErrorAction Stop
            Write-Host "Started service: $svc" -ForegroundColor Green
            $log += [PSCustomObject]@{ Step = "Start Service"; Detail = $svc; Status = "Started" }
        } catch {
            Write-Host "Failed to start service: $svc" -ForegroundColor Yellow
            $log += [PSCustomObject]@{ Step = "Start Service"; Detail = $_.Exception.Message; Status = "Failed" }
        }
    }

    $log | Export-Csv -Path $csvLogFile -NoTypeInformation -Encoding UTF8
    Write-Host "`nFull output saved to: $transcriptFile" -ForegroundColor Cyan
    Write-Host "Summary CSV saved to: $csvLogFile" -ForegroundColor Cyan

    Stop-Transcript
    Write-Host "`n=== AgentCheck Task Completed ===" -ForegroundColor Magenta
}

function Run-AgentInstallUtility {
    $exportDir = "C:\Script-Export"
    if (-not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $hostname = $env:COMPUTERNAME
    $csvLogFile = "$exportDir\AgentMaintenance-$timestamp-$hostname.csv"
    $textLogFile = "$exportDir\AgentMaintenance-FullOutput-$timestamp-$hostname.txt"
    Start-Transcript -Path $textLogFile -Append -Force
    $log = @()

    function Write-Log { param ([string]$action,[string]$target,[string]$result)
        $global:log += [PSCustomObject]@{
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Action    = $action; Target = $target; Result = $result }
    }

    $agentPath = "C:\Program Files (x86)\CyberCNSAgent\uninstall.bat"
    $services = @("cybercnsagent", "cybercnsagentmonitor")

    if (Test-Path $agentPath) {
        $uninstallScript = Get-Content $agentPath -Raw
        $backup = "$exportDir\UninstallScriptBackup-$timestamp-$hostname.txt"
        $uninstallScript | Out-File -FilePath $backup -Encoding UTF8
        Write-Log "ExportFile" "uninstall.bat Backup" "Saved"
        cmd /c `"$agentPath`"
        Write-Log "Uninstall" "Agent" "Attempted"
    }

    $companyId = Read-Host "Enter Company ID"
    $tenantId  = Read-Host "Enter Tenant ID"
    $secretKey = Read-Host "Enter Secret Key"
    $command = ".\cybercnsagent.exe -c $companyId -e $tenantId -j $secretKey -i"
    Write-Host "Command: $command" -ForegroundColor Yellow
    $confirm = Read-Host "Run install command? (Y/N)"
    if ($confirm -match '^[Yy]$') {
        if (-not (Test-Path ".\cybercnsagent.exe")) {
            Invoke-WebRequest -Uri "https://agentv3.myconnectsecure.com/cybercnsagent.exe" -OutFile ".\cybercnsagent.exe" -UseBasicParsing
            Write-Log "Download" "cybercnsagent.exe" "Success"
        }
        cmd /c $command
        Write-Log "Install" "Agent" "Executed"
    } else {
        Write-Host "Installation cancelled."
        Write-Log "Install" "Agent" "Cancelled"
    }

    $log | Export-Csv -Path $csvLogFile -NoTypeInformation -Encoding UTF8
    Stop-Transcript
    Write-Host "Logs saved to: $csvLogFile and $textLogFile"
}

function Run-AgentCheckSMB {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputDir = "C:\Script-Export"
    if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }
    $outputFile = "$outputDir\SMB_Version_Report_$timestamp.csv"
    $SmbRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
    $results = @()

    $Smb1Enabled = Get-ItemProperty -Path $SmbRegistryPath -Name "SMB1" -ErrorAction SilentlyContinue
    $results += [PSCustomObject]@{
        SMBVersion = "SMB 1.0"
        Status = if ($Smb1Enabled.SMB1 -eq 1) {"Enabled"} else {"Disabled"}
        DetectionKey = "$SmbRegistryPath\SMB1"
        DisableKey = if ($Smb1Enabled.SMB1 -eq 1) {"$SmbRegistryPath\SMB1"} else {""}
    }

    $Smb2Enabled = Get-ItemProperty -Path $SmbRegistryPath -Name "SMB2" -ErrorAction SilentlyContinue
    $results += [PSCustomObject]@{
        SMBVersion = "SMB 2.0"
        Status = if ($Smb2Enabled.SMB2 -eq 1) {"Enabled"} else {"Disabled"}
        DetectionKey = "$SmbRegistryPath\SMB2"
        DisableKey = if ($Smb2Enabled.SMB2 -eq 1) {"$SmbRegistryPath\SMB2"} else {""}
    }

    $Smb3Enabled = Get-ItemProperty -Path $SmbRegistryPath -Name "SMB3" -ErrorAction SilentlyContinue
    $results += [PSCustomObject]@{
        SMBVersion = "SMB 3.0"
        Status = if ($Smb3Enabled.SMB3 -eq 1) {"Enabled"} else {"Disabled"}
        DetectionKey = "$SmbRegistryPath\SMB3"
        DisableKey = if ($Smb3Enabled.SMB3 -eq 1) {"$SmbRegistryPath\SMB3"} else {""}
    }

    $results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "Report saved to: $outputFile" -ForegroundColor Green
}

function Run-AgentSetSMB {
    Write-Host "`n=== Enabling SMB Functionality ===" -ForegroundColor Cyan
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name SMB2 -Type DWORD -Value 1 -Force
        Write-Host "✔ SMB2 protocol enabled in registry." -ForegroundColor Green
    } catch {
        Write-Host "✖ Failed to enable SMB2 in registry: $_" -ForegroundColor Red
    }

    $firewallRules = @("File And Printer Sharing (SMB-In)", "File And Printer Sharing (NB-Session-In)")
    foreach ($rule in $firewallRules) {
        try {
            Set-NetFirewallRule -DisplayName $rule -Enabled True -Profile Any
            Write-Host "✔ Firewall rule '$rule' enabled." -ForegroundColor Green
        } catch {
            Write-Host "✖ Failed to enable firewall rule '$rule': $_" -ForegroundColor Red
        }
    }

    try {
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
            -Name LocalAccountTokenFilterPolicy -PropertyType DWord -Value 1 -Force | Out-Null
        Write-Host "✔ LocalAccountTokenFilterPolicy registry key set." -ForegroundColor Green
    } catch {
        Write-Host "✖ Failed to set LocalAccountTokenFilterPolicy: $_" -ForegroundColor Red
    }
    Write-Host "`n=== SMB Function Enablement Complete ===" -ForegroundColor Cyan
}

function Run-SelectedOption {
    param($choice)
    switch ($choice.ToUpper()) {
        "1" { Run-ValidationScripts }
        "2" { Run-AgentMaintenance }
        "3" { Write-Host "[Placeholder] Probe Troubleshooting" }
        "4" { Run-ZipAndEmailResults }
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
