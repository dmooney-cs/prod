<# =================================================================================================
 CS-Toolbox-Download-Verify-Extract-NO-LAUNCH.ps1  (v2.3 - GitHub URL hardening + staging extract)

 What this does:
  - Downloads a ZIP (supports GitHub "github.com/.../raw/..." AND raw.githubusercontent.com URLs)
  - Downloads SHA-256 text file and parses the first token as the hash
  - Verifies ZIP SHA-256 (unless -SkipHashCheck)
  - Extracts into C:\CS-Toolbox-TEMP\prod-01-01 using isolated staging + safe normalize
  - Overwrites existing files (-Force) and never creates "(2)" copies
  - DOES NOT LAUNCH anything
  - Exits 0 on success, 1 on failure (no "True" output)

 Switches:
  -SkipHashCheck
  -ShowHashes
  -ConfirmHashes
  -ExportOnly  : exports JSON to C:\Temp\collected-info and exits (no prompts unless -ConfirmHashes)
================================================================================================= #>

#requires -version 5.1
[CmdletBinding()]
param(
    [switch]$SkipHashCheck = $false,
    [switch]$ShowHashes    = $false,
    [switch]$ConfirmHashes = $false,
    [switch]$ExportOnly    = $false
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --------------------------
# Config (YOUR DIFFERENT ZIP + SHA FILE)
# --------------------------
$ZipUrl      = 'https://github.com/dmooney-cs/prod/raw/refs/heads/main/prod-01-01.zip'
$ZipPath     = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath = 'C:\CS-Toolbox-TEMP'
$DestRoot    = Join-Path $ExtractPath 'prod-01-01'

$HashUrl     = 'https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/devprod.sha256'

# --------------------------
# Output control
# --------------------------
$script:AllowOutput = [bool]($ShowHashes -or $ConfirmHashes)

function Say([string]$msg) {
    if ($script:AllowOutput) { Write-Host $msg }
}

function Export-ResultJson {
    param(
        [hashtable]$Obj
    )
    try {
        $dir = 'C:\Temp\collected-info'
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        $name = "CS-Toolbox-Download-Verify-Extract-NO-LAUNCH_{0}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss")
        $path = Join-Path $dir $name
        ($Obj | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $path -Encoding UTF8
    } catch { }
}

function Fail([string]$stage, [string]$reason) {
    if ($script:AllowOutput) {
        Write-Host ""
        Write-Host "FAILED at: $stage"
        Write-Host "Reason : $reason"
        Write-Host ""
    }

    if ($ExportOnly) {
        Export-ResultJson @{
            ok        = $false
            stage     = $stage
            reason    = $reason
            zipUrl    = $ZipUrl
            hashUrl   = $HashUrl
            zipPath   = $ZipPath
            destRoot  = $DestRoot
            ts        = (Get-Date).ToString("o")
        }
    }

    exit 1
}

# --------------------------
# URL helpers (fix common GitHub URL mistakes)
# --------------------------
function Convert-GitHubToRawUrl {
    param([Parameter(Mandatory)][string]$Uri)

    # Handle:
    #  - https://github.com/OWNER/REPO/raw/refs/heads/BRANCH/path
    #  - https://github.com/OWNER/REPO/blob/BRANCH/path
    # Convert to:
    #  - https://raw.githubusercontent.com/OWNER/REPO/refs/heads/BRANCH/path
    $u = $Uri.Trim()

    if ($u -match '^https://github\.com/([^/]+)/([^/]+)/raw/(.+)$') {
        return ('https://raw.githubusercontent.com/{0}/{1}/{2}' -f $Matches[1], $Matches[2], $Matches[3])
    }
    if ($u -match '^https://github\.com/([^/]+)/([^/]+)/blob/(.+)$') {
        return ('https://raw.githubusercontent.com/{0}/{1}/{2}' -f $Matches[1], $Matches[2], $Matches[3])
    }
    return $u
}

function File-LooksLikeHtml {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $false }
        $len = (Get-Item -LiteralPath $Path).Length
        if ($len -lt 20) { return $false }
        $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 512
        $text  = [Text.Encoding]::UTF8.GetString($bytes)
        return ($text -match '<!DOCTYPE html' -or $text -match '<html' -or $text -match 'github' -and $text -match '<title>')
    } catch {
        return $false
    }
}

# --------------------------
# Download helpers
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
                'User-Agent'    = 'CS-Toolbox'
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

function Invoke-DownloadWithGitHubFix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$OutFile,
        [int]$Attempts = 3
    )

    # 1) Try as provided
    if (Invoke-DownloadQuiet -Uri $Uri -OutFile $OutFile -Attempts $Attempts) {
        # If GitHub served HTML, try converting to raw
        if (-not (File-LooksLikeHtml -Path $OutFile)) { return $true }
    }

    # 2) Try converted raw URL (handles github.com/blob and github.com/raw forms)
    $raw = Convert-GitHubToRawUrl -Uri $Uri
    if ($raw -ne $Uri) {
        try { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue } catch { }
        if (Invoke-DownloadQuiet -Uri $raw -OutFile $OutFile -Attempts $Attempts) {
            if (-not (File-LooksLikeHtml -Path $OutFile)) { return $true }
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

    $u = Convert-GitHubToRawUrl -Uri $Uri

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $r = Invoke-WebRequest -Uri $u -UseBasicParsing -ErrorAction Stop -Headers @{
                'Cache-Control' = 'no-cache'
                'Pragma'        = 'no-cache'
                'User-Agent'    = 'CS-Toolbox'
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

# --------------------------
# Extract helpers (staging + safe normalize)
# --------------------------
function New-IsolatedStageFolder {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Parent)

    $name = "_stage_" + ([guid]::NewGuid().ToString("N"))
    $path = Join-Path $Parent $name
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

function Resolve-ExtractedRoot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$StageRoot)

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

    foreach ($item in (Get-ChildItem -LiteralPath $ContentRoot -Force -ErrorAction SilentlyContinue)) {
        try {
            Move-Item -LiteralPath $item.FullName -Destination $DestRoot -Force -ErrorAction Stop
        } catch {
            try { Copy-Item -LiteralPath $item.FullName -Destination $DestRoot -Recurse -Force -ErrorAction Stop } catch { }
            try { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch { }
        }
    }
}

# --------------------------
# Main
# --------------------------
$stage = "init"
$expected = $null
$actual   = $null
$finalZipUrlUsed = $ZipUrl
$finalHashUrlUsed = $HashUrl

try {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    $stage = "prepare folders"
    New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
    Remove-Item -LiteralPath $DestRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -Path $DestRoot -ItemType Directory -Force | Out-Null

    $stage = "download zip"
    Say "Downloading ZIP..."
    if (-not (Invoke-DownloadWithGitHubFix -Uri $ZipUrl -OutFile $ZipPath -Attempts 3)) {
        Fail $stage "Unable to download ZIP from $ZipUrl (or converted raw URL)"
    }

    if (File-LooksLikeHtml -Path $ZipPath) {
        Fail $stage "Downloaded content looks like HTML (likely wrong GitHub URL / auth / redirect)."
    }

    if (-not $SkipHashCheck) {
        $stage = "download expected hash"
        Say "Fetching expected SHA-256..."
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
            Write-Host "Expected (SHA file): $expected"
            Write-Host "Actual   (ZIP)     : $actual"
            Write-Host ""
        }

        $stage = "compare hashes"
        if ($actual -ne $expected) {
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            Fail $stage "Hash mismatch"
        }

        if ($ConfirmHashes) {
            $resp = Read-Host "Hashes match. Proceed with extract? (Y/N)"
            if ($resp -notin @('Y','y')) {
                if ($ExportOnly) {
                    Export-ResultJson @{
                        ok        = $true
                        stage     = "user_cancel"
                        zipUrl    = $ZipUrl
                        hashUrl   = $HashUrl
                        expected  = $expected
                        actual    = $actual
                        zipPath   = $ZipPath
                        destRoot  = $DestRoot
                        ts        = (Get-Date).ToString("o")
                    }
                }
                exit 0
            }
        }
    } else {
        if ($ConfirmHashes -or $ShowHashes) {
            Write-Host ""
            Write-Host "Hash check skipped (-SkipHashCheck)."
            Write-Host ""
        }
        if ($ConfirmHashes) {
            $resp = Read-Host "Proceed without hash verification? (Y/N)"
            if ($resp -notin @('Y','y')) {
                if ($ExportOnly) {
                    Export-ResultJson @{
                        ok        = $true
                        stage     = "user_cancel"
                        zipUrl    = $ZipUrl
                        hashUrl   = $HashUrl
                        expected  = $null
                        actual    = $null
                        zipPath   = $ZipPath
                        destRoot  = $DestRoot
                        ts        = (Get-Date).ToString("o")
                    }
                }
                exit 0
            }
        }
    }

    $stage = "extract zip (staging)"
    Say "Extracting (staging)..."
    $stageRoot = $null
    try {
        $stageRoot = New-IsolatedStageFolder -Parent $ExtractPath
        Expand-Archive -Path $ZipPath -DestinationPath $stageRoot -Force
    } catch {
        if ($stageRoot) { try { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { } }
        Fail $stage $_.Exception.Message
    }

    $stage = "resolve extracted root"
    $contentRoot = Resolve-ExtractedRoot -StageRoot $stageRoot

    $stage = "move content into destination"
    Move-TreeIntoDest -ContentRoot $contentRoot -DestRoot $DestRoot

    $stage = "cleanup staging"
    try { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

    $stage = "unblock"
    try {
        Get-ChildItem -LiteralPath $DestRoot -Recurse -Force -File -ErrorAction SilentlyContinue |
            ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch { } }
    } catch { }

    if ($ExportOnly) {
        Export-ResultJson @{
            ok        = $true
            stage     = "done"
            zipUrl    = $ZipUrl
            hashUrl   = $HashUrl
            expected  = $expected
            actual    = $actual
            zipPath   = $ZipPath
            destRoot  = $DestRoot
            ts        = (Get-Date).ToString("o")
        }
    }

    # DONE (no launch) - no "True" output
    exit 0

} catch {
    Fail $stage $_.Exception.Message
}
