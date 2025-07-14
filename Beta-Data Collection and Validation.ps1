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

function Start-Tool {
    do {
        Show-MainMenu
        $choice = Read-Host "Enter your choice"
        Run-SelectedOption -choice $choice
    } while ($choice.ToUpper() -ne "Q")
}

Start-Tool
