```powershell
<# =================================================================================================
 CS-Toolbox-Launcher-DevTools-ZeroTouch-NO-LAUNCH.ps1  (v2.2b - NO-EXIT fix)

 Key fix from v2.2a:
  - Replaced ALL `exit` usage with `return` / `throw`
    so running via `irm ... | iex` does NOT terminate the caller session.
================================================================================================= #>

#requires -version 5.1
[CmdletBinding()]
param(
    [switch]$SkipHashCheck = $false,
    [switch]$ShowHashes = $false,
    [switch]$ConfirmHashes = $false,
    [switch]$ExportOnly = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --------------------------
# Config
# --------------------------
$ZipUrl      = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath     = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath = 'C:\CS-Toolbox-TEMP'
$DestRoot    = Join-Path $ExtractPath 'prod-01-01'
$Launcher    = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1' # retained only for validation parity (not launched)

# IMPORTANT: GitHub RAW hash file URL
$HashUrl     = 'https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/devtools.sha256'

# --------------------------
# Output control
# --------------------------
# Silent unless hashes are involved (or you choose to change this behavior)
$script:AllowOutput = [bool]($ShowHashes -or $ConfirmHashes)

function Say([string]$msg) {
    if ($script:AllowOutput) { Write-Host $msg }
}

function Fail([string]$stage, [string]$reason) {
    if ($script:AllowOutput) {
        Write-Host ""
        Write-Host "FAILED at: $stage"
        Write-Host "Reason : $reason"
        Write-Host ""
    }
    # IMPORTANT: do NOT exit the host session (esp. when invoked via irm|iex)
    throw "FAILED at: $stage | $reason"
}

# --------------------------
# Helpers (quiet)
# --------------------------
function Invoke-DownloadQuiet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$Attempts = 3
    )

    try { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue } catch { }

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }
            if ((Test-Path -LiteralPath $OutFile) -and ((Get-Item -LiteralPath $OutFile).Length -gt 0)) {
                return $true
            }
        } catch {
            try { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue } catch { }
            Start-Sleep -Seconds 1
        }
    }
    return $false
}

function Get-RemoteSha256Quiet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$Attempts = 3
    )

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $Uri -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
            }

            $txt = ($r.Content | Out-String)
            if ([string]::IsNullOrWhiteSpace($txt)) { throw "empty response" }

            $line = ($txt -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1).Trim()
            if ([string]::IsNullOrWhiteSpace($line)) { throw "blank file" }

            $hash = ($line -split '\s+')[0].Trim().ToUpperInvariant()
            if ($hash -match '^[0-9A-F]{64}$') { return $hash }

            throw "not a valid SHA-256 (expected 64 hex chars)"
        } catch {
            Start-Sleep -Seconds 1
        }
    }

    return $null
}

function New-IsolatedStageFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Parent
    )
    $name = "_stage_" + ([guid]::NewGuid().ToString("N"))
    $path = Join-Path $Parent $name
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

function Resolve-ExtractedRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$StageRoot
    )

    # If exactly one top-level folder exists and no top-level files exist,
    # treat that folder as the content root (flatten one wrapper folder).
    $topDirs  = @(Get-ChildItem -LiteralPath $StageRoot -Directory -Force -ErrorAction SilentlyContinue)
    $topFiles = @(Get-ChildItem -LiteralPath $StageRoot -File      -Force -ErrorAction SilentlyContinue)

    if ($topDirs.Count -eq 1 -and $topFiles.Count -eq 0) {
        return $topDirs[0].FullName
    }
    return $StageRoot
}

function Move-TreeIntoDest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ContentRoot,
        [Parameter(Mandatory)][string]$DestRoot
    )

    # Move everything under ContentRoot into DestRoot, preserving structure
    # Use -Force to overwrite existing files (no "(2)" copies)
    foreach ($item in (Get-ChildItem -LiteralPath $ContentRoot -Force -ErrorAction SilentlyContinue)) {
        try {
            Move-Item -LiteralPath $item.FullName -Destination $DestRoot -Force -ErrorAction Stop
        } catch {
            # Fallback: copy+delete best effort (handles some locks/ACL oddities)
            try { Copy-Item -LiteralPath $item.FullName -Destination $DestRoot -Recurse -Force -ErrorAction Stop } catch { }
            try { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
}

# --------------------------
# Main
# --------------------------
$stage = "init"

try {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    $stage = "prepare folders"
    try { New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null } catch { Fail $stage $_.Exception.Message }

    # Ensure a clean destination root
    try { Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    try { New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null } catch { Fail $stage $_.Exception.Message }

    $stage = "download zip"
    Say "Downloading ZIP..."
    if (-not (Invoke-DownloadQuiet -Uri $ZipUrl -OutFile $ZipPath -Attempts 3)) {
        Fail $stage "Unable to download ZIP from $ZipUrl"
    }

    $expected = $null
    $actual   = $null

    if (-not $SkipHashCheck) {
        $stage = "download expected hash"
        Say "Fetching expected SHA-256 from GitHub..."
        $expected = Get-RemoteSha256Quiet -Uri $HashUrl -Attempts 3
        if ([string]::IsNullOrWhiteSpace($expected)) {
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            Fail $stage "Unable to download/parse SHA-256 from $HashUrl"
        }

        $stage = "compute zip hash"
        Say "Computing ZIP SHA-256..."
        try {
            $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpperInvariant()
        } catch {
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            Fail $stage $_.Exception.Message
        }

        if ($ShowHashes -or $ConfirmHashes) {
            Write-Host ""
            Write-Host "Expected (GitHub): $expected"
            Write-Host "Actual   (ZIP)   : $actual"
            Write-Host ""
        }

        $stage = "compare hashes"
        if ($actual -ne $expected) {
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            Fail $stage "Hash mismatch"
        }

        if ($ConfirmHashes) {
            $resp = Read-Host "Hashes match. Proceed with extract? (Y/N)"
            if ($resp -notin @('Y','y')) { return }
        }
    } else {
        if ($ConfirmHashes -or $ShowHashes) {
            Write-Host ""
            Write-Host "Hash check skipped (-SkipHashCheck)."
            Write-Host ""
        }
        if ($ConfirmHashes) {
            $resp = Read-Host "Proceed without hash verification? (Y/N)"
            if ($resp -notin @('Y','y')) { return }
        }
    }

    $stage = "extract zip (staging)"
    Say "Extracting (staging)..."
    $StageRoot = $null
    try {
        $StageRoot = New-IsolatedStageFolder -Parent $ExtractPath
        Expand-Archive -Path $ZipPath -DestinationPath $StageRoot -Force
    } catch {
        if ($StageRoot) { try { Remove-Item -LiteralPath $StageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { } }
        Fail $stage $_.Exception.Message
    }

    $stage = "resolve extracted root"
    $ContentRoot = $null
    try {
        $ContentRoot = Resolve-ExtractedRoot -StageRoot $StageRoot
    } catch {
        try { Remove-Item -LiteralPath $StageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        Fail $stage $_.Exception.Message
    }

    $stage = "move content into destination"
    try {
        Move-TreeIntoDest -ContentRoot $ContentRoot -DestRoot $DestRoot
    } catch {
        try { Remove-Item -LiteralPath $StageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        Fail $stage $_.Exception.Message
    }

    $stage = "cleanup staging"
    try { Remove-Item -LiteralPath $StageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

    $stage = "unblock"
    try {
        Get-ChildItem -LiteralPath $DestRoot -Recurse -Force -File -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { }
            }
    } catch { }

    # --------------------------
    # DONE (no launch) - IMPORTANT: do NOT exit the host session
    # --------------------------
    return

} catch {
    Fail $stage $_.Exception.Message
}
```
