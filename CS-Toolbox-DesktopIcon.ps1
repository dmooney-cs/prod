<# =================================================================================================
 CS-Toolbox-DesktopIcon.ps1  (v2.0 - overwrite always + include .ico + default minimal UI + ShowDetails)

 Default behavior:
  - Runs without showing the full planned-actions menu
  - After successful extraction + copy, displays:
      "Please find the ConnectSecure Toolbox Launcher icon on your desktop, press any key to exit"
    then waits for any key.

 -ShowDetails:
  - Shows the detailed planned actions + lists .ico files to be copied
  - Prompts to proceed (Y/N)

 Other:
  - NEVER create "_(2)" copies. ALWAYS overwrite using -Force.
  - Copies ALL .ps1 and ALL .ico to C:\Temp (overwrite).
  - Copies ALL .lnk to Desktop and/or Taskbar pinned folder (overwrite).
  - -ExportOnly exports JSON to C:\Temp\collected-info and exits (no prompts).

================================================================================================= #>

#requires -version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$ZipUrl = "https://github.com/dmooney-cs/prod/raw/refs/heads/main/Toolbox-Launcher.zip",

    [switch]$Desktop,
    [switch]$Taskbar,

    [switch]$ShowDetails,
    [switch]$Silent,

    [switch]$ExportOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if (-not $Desktop -and -not $Taskbar) { $Desktop = $true }

# ------------------------- Paths -------------------------
$DeployRoot       = "C:\Temp"
$CollectedInfoDir = Join-Path $DeployRoot "collected-info"
$LogFile          = Join-Path $DeployRoot "CS-Toolbox-DesktopIcon.log"
$ExportJson       = Join-Path $CollectedInfoDir "CS-Toolbox-DesktopIcon.json"

function Ensure-Dir {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}
Ensure-Dir $DeployRoot
Ensure-Dir $CollectedInfoDir

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
    if (-not $Silent -and $ShowDetails) {
        switch ($Level) {
            'ERROR' { Write-Host $line -ForegroundColor Red }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'OK'    { Write-Host $line -ForegroundColor Green }
            default { Write-Host $line }
        }
    }
}
$script:LogFile = $LogFile

function Get-LoggedOnUserSid {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    if (-not $cs.UserName) { throw "No interactive user detected (Win32_ComputerSystem.UserName is empty)." }
    $nt  = New-Object System.Security.Principal.NTAccount($cs.UserName)
    $sid = $nt.Translate([System.Security.Principal.SecurityIdentifier]).Value
    [pscustomobject]@{ UserName = $cs.UserName; Sid = $sid }
}

function Get-UserShellFolderPath {
    param(
        [Parameter(Mandatory)][string]$Sid,
        [Parameter(Mandatory)][ValidateSet('Desktop','AppData')][string]$Folder
    )
    $base = "Registry::HKEY_USERS\$Sid\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    $name = if ($Folder -eq 'Desktop') { 'Desktop' } else { 'AppData' }
    $val  = (Get-ItemProperty -Path $base -Name $name -ErrorAction Stop).$name
    if (-not $val) { throw "Unable to resolve $Folder path from HKU:\$Sid Shell Folders." }
    $val
}

function Get-FileHashSafe {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Invoke-DownloadWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )

    if (Test-Path -LiteralPath $OutFile) {
        Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Write-Log ("Downloading... Attempt {0}/{1}" -f $attempt, $MaxAttempts) "INFO"
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop
            if ((Test-Path -LiteralPath $OutFile) -and ((Get-Item -LiteralPath $OutFile).Length -gt 0)) {
                Write-Log "Download successful." "OK"
                return $true
            }
            throw "Downloaded file is missing or empty."
        } catch {
            Write-Log ("Attempt {0}/{1} failed: {2}" -f $attempt, $MaxAttempts, $_.Exception.Message) "WARN"
            if (Test-Path -LiteralPath $OutFile) {
                Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
            }
            if ($attempt -eq $MaxAttempts) { return $false }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
    return $false
}

function Move-Contents {
    param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$Target)
    if (-not (Test-Path -LiteralPath $Source)) { return }
    Get-ChildItem -LiteralPath $Source -Force | ForEach-Object {
        try { Move-Item -LiteralPath $_.FullName -Destination $Target -Force } catch {
            Write-Log ("Failed to move {0} -> {1}: {2}" -f $_.FullName, $Target, $_.Exception.Message) "WARN"
        }
    }
}

function Normalize-ExtractRoot {
    param(
        [Parameter(Mandatory)][string]$ExtractPath,
        [Parameter(Mandatory)][string]$ZipPath
    )

    $topDirs  = @(
        Get-ChildItem -LiteralPath $ExtractPath -Directory -Force -ErrorAction SilentlyContinue
    )
    $topFiles = @(
        Get-ChildItem -LiteralPath $ExtractPath -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -ne $ZipPath }
    )

    if ($topDirs.Count -eq 1 -and $topFiles.Count -eq 0) {
        Write-Log ("Normalizing: flattening single folder {0}" -f $topDirs[0].Name) "INFO"
        Move-Contents -Source $topDirs[0].FullName -Target $ExtractPath
        try { Remove-Item -LiteralPath $topDirs[0].FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    } else {
        Write-Log ("Normalize skipped: topDirs={0} topFiles={1}" -f $topDirs.Count, $topFiles.Count) "INFO"
    }
}

function Unblock-AllFiles {
    param([Parameter(Mandatory)][string]$Root)
    try {
        Get-ChildItem -LiteralPath $Root -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object {
            try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { }
        }
    } catch { }
}

function Get-AllFilesByExtension {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Extension)
    Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction Stop |
        Where-Object { $_.Extension -ieq $Extension }
}

function Copy-AllFilesOverwrite {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)][string]$Destination
    )
    Ensure-Dir $Destination

    $copied = @()
    foreach ($f in $Files) {
        $dest = Join-Path $Destination $f.Name
        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
        $copied += $dest
        Write-Log "Copied (overwrite): $($f.FullName) -> $dest" "OK"
    }
    return $copied
}

function Wait-AnyKey {
    try {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } catch {
        Start-Sleep -Seconds 2
    }
}

# ------------------------- Summary -------------------------
$summary = [ordered]@{
    startedAt           = (Get-Date).ToString('o')
    zipUrl              = $ZipUrl
    downloadedZip       = $null
    downloadSha256      = $null
    extractTo           = $null
    interactiveUser     = $null
    userSid             = $null
    userDesktop         = $null
    userAppData         = $null
    taskbarPinnedFolder = $null
    lnkCountFound       = 0
    ps1CountFound       = 0
    icoCountFound       = 0
    copiedLnkToDesktop  = @()
    copiedLnkToTaskbar  = @()
    copiedPs1ToCTemp    = @()
    copiedIcoToCTemp    = @()
    result              = "UNKNOWN"
}

Write-Log "Starting. ZipUrl='$ZipUrl' Desktop=$Desktop Taskbar=$Taskbar ShowDetails=$ShowDetails Silent=$Silent ExportOnly=$ExportOnly" "INFO"

try {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    $user = Get-LoggedOnUserSid
    $summary.interactiveUser = $user.UserName
    $summary.userSid = $user.Sid
    Write-Log "Interactive user: $($user.UserName) SID=$($user.Sid)" "INFO"

    $desktopPath = Get-UserShellFolderPath -Sid $user.Sid -Folder Desktop
    $appDataPath = Get-UserShellFolderPath -Sid $user.Sid -Folder AppData
    $summary.userDesktop = $desktopPath
    $summary.userAppData = $appDataPath

    $taskbarDir = Join-Path $appDataPath "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $summary.taskbarPinnedFolder = $taskbarDir

    $sysTemp = Join-Path $env:windir "Temp"
    Ensure-Dir $sysTemp

    $stamp   = (Get-Date -Format "yyyyMMdd_HHmmss")
    $shortId = ([guid]::NewGuid().ToString("N").Substring(0,8))
    $zipPath = Join-Path $sysTemp ("Toolbox-Launchers_{0}_{1}.zip" -f $stamp, $shortId)
    $extract = Join-Path $sysTemp ("CS-Toolbox-Extract_{0}_{1}" -f $stamp, $shortId)

    $ok = Invoke-DownloadWithRetry -Uri $ZipUrl -OutFile $zipPath -MaxAttempts 3 -DelaySeconds 2
    if (-not $ok) { throw "Download did not succeed after 3 attempts." }

    $summary.downloadedZip  = $zipPath
    $summary.downloadSha256 = Get-FileHashSafe -Path $zipPath
    Write-Log "Downloaded ZIP SHA256: $($summary.downloadSha256)" "OK"

    if (Test-Path -LiteralPath $extract) { Remove-Item -LiteralPath $extract -Recurse -Force -ErrorAction SilentlyContinue }
    Ensure-Dir $extract

    Write-Log "Extracting to: $extract" "INFO"
    Expand-Archive -Path $zipPath -DestinationPath $extract -Force
    Write-Log "Extract complete." "OK"

    $summary.extractTo = $extract

    Normalize-ExtractRoot -ExtractPath $extract -ZipPath $zipPath
    Unblock-AllFiles -Root $extract

    $lnkFiles = @(Get-AllFilesByExtension -Root $extract -Extension ".lnk")
    $ps1Files = @(Get-AllFilesByExtension -Root $extract -Extension ".ps1")
    $icoFiles = @(Get-AllFilesByExtension -Root $extract -Extension ".ico")

    $summary.lnkCountFound = $lnkFiles.Count
    $summary.ps1CountFound = $ps1Files.Count
    $summary.icoCountFound = $icoFiles.Count

    Write-Log "Found .lnk files: $($lnkFiles.Count)" "OK"
    Write-Log "Found .ps1 files: $($ps1Files.Count)" "OK"
    Write-Log "Found .ico files: $($icoFiles.Count)" "OK"

    if ($ExportOnly) {
        $summary.result = "EXPORTONLY"
        $summary.finishedAt = (Get-Date).ToString('o')
        ($summary | ConvertTo-Json -Depth 14) | Set-Content -Path $ExportJson -Encoding UTF8
        exit 0
    }

    # Show details only when requested
    if (-not $Silent -and $ShowDetails) {
        Write-Host ""
        Write-Host "Planned actions (OVERWRITE existing files):" -ForegroundColor Cyan
        Write-Host (" - Copy {0} .ps1 -> {1}" -f $ps1Files.Count, $DeployRoot)
        Write-Host (" - Copy {0} .ico -> {1}" -f $icoFiles.Count, $DeployRoot)

        if ($icoFiles.Count -gt 0) {
            Write-Host "   ICO files to be copied:" -ForegroundColor Cyan
            foreach ($ico in $icoFiles) { Write-Host ("    • {0}" -f $ico.FullName) }
        }

        if ($Desktop) { Write-Host (" - Copy {0} .lnk -> Desktop: {1}" -f $lnkFiles.Count, $desktopPath) }
        if ($Taskbar) { Write-Host (" - Copy {0} .lnk -> Taskbar pinned folder: {1}" -f $lnkFiles.Count, $taskbarDir) }

        Write-Host ""
        $ans = Read-Host "Proceed? (Y/N)"
        if ($ans -notin @('Y','y')) { throw "User cancelled." }
    }

    # Copy (overwrite)
    if ($ps1Files.Count -gt 0) { $summary.copiedPs1ToCTemp = Copy-AllFilesOverwrite -Files $ps1Files -Destination $DeployRoot }
    if ($icoFiles.Count -gt 0) { $summary.copiedIcoToCTemp = Copy-AllFilesOverwrite -Files $icoFiles -Destination $DeployRoot }

    if ($Desktop -and $lnkFiles.Count -gt 0) {
        $summary.copiedLnkToDesktop = Copy-AllFilesOverwrite -Files $lnkFiles -Destination $desktopPath
    }
    if ($Taskbar -and $lnkFiles.Count -gt 0) {
        Ensure-Dir $taskbarDir
        $summary.copiedLnkToTaskbar = Copy-AllFilesOverwrite -Files $lnkFiles -Destination $taskbarDir
    }

    $summary.result = "SUCCESS"

    if (-not $Silent) {
        Write-Host ""
        Write-Host "Please find the ConnectSecure Toolbox Launcher icon on your desktop, press any key to exit"
        Wait-AnyKey
    }
}
catch {
    $summary.result = "FAILED"
    $summary.error  = $_.Exception.Message
    Write-Log "FAILED: $($summary.error)" "ERROR"
    if (-not $Silent) { throw }
}
finally {
    $summary.finishedAt = (Get-Date).ToString('o')
    ($summary | ConvertTo-Json -Depth 14) | Set-Content -Path $ExportJson -Encoding UTF8
    Write-Log "Summary exported to: $ExportJson" "INFO"
    Write-Log "Done. Result=$($summary.result)" "INFO"
}

exit 0
