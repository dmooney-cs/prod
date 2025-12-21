# CS-Toolbox-Launcher-ZeroTouch-SILENT.ps1
# Fully silent bootstrapper (NO output, NO prompts, NO progress)

param(
    [switch]$SkipHashCheck = $false
)

#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# --------------------------
# Config
# --------------------------
$ZipUrl       = 'https://betadevtools.myconnectsecure.com/agents/toolbox/toolbox.zip'
$ZipPath      = Join-Path $env:TEMP 'prod-01-01.zip'
$ExtractPath  = 'C:\CS-Toolbox-TEMP'
$DestRoot     = Join-Path $ExtractPath 'prod-01-01'
$Launcher     = Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'

$ExpectedHash = '383A1C3E9AABD95365572C67D64CDEF51604A967561D79EBD299FB877E6352C0'

# --------------------------
# Helpers (silent)
# --------------------------
function Invoke-DownloadSilent {
    param(
        [string]$Uri,
        [string]$OutFile,
        [int]$Attempts = 3
    )

    Remove-Item $OutFile -Force -ErrorAction SilentlyContinue

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Invoke-WebRequest `
                -Uri $Uri `
                -OutFile $OutFile `
                -UseBasicParsing `
                -ErrorAction Stop `
                -Headers @{ 'Cache-Control' = 'no-cache' }

            if ((Get-Item $OutFile).Length -gt 0) {
                return $true
            }
        } catch {
            Remove-Item $OutFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
        }
    }
    return $false
}

function Move-ContentsSilent {
    param($Source, $Target)
    if (-not (Test-Path $Source)) { return }
    Get-ChildItem $Source -Force | ForEach-Object {
        try { Move-Item $_.FullName $Target -Force -ErrorAction Stop } catch { }
    }
}

# --------------------------
# Execution (ALL SILENT)
# --------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

try { New-Item $ExtractPath -ItemType Directory -Force | Out-Null } catch { return }
try { Remove-Item $DestRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

# Download
if (-not (Invoke-DownloadSilent -Uri $ZipUrl -OutFile $ZipPath)) { return }

# Hash verification
if (-not $SkipHashCheck) {
    try {
        $actual = (Get-FileHash $ZipPath -Algorithm SHA256).Hash.ToUpper()
        if ($actual -ne $ExpectedHash) {
            Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
            return
        }
    } catch {
        Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
        return
    }
}

# Extract
try { Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force } catch { return }

# Ensure destination
try { New-Item $DestRoot -ItemType Directory -Force | Out-Null } catch { return }

# Normalize ZIP structure
if (-not (Test-Path (Join-Path $DestRoot 'CS-Toolbox-Launcher.ps1'))) {
    $dirs  = Get-ChildItem $ExtractPath -Directory -Force | Where-Object { $_.FullName -ne $DestRoot }
    $files = Get-ChildItem $ExtractPath -File -Force | Where-Object { $_.FullName -ne $ZipPath }

    if ($dirs.Count -eq 1 -and $files.Count -eq 0) {
        Move-ContentsSilent $dirs[0].FullName $DestRoot
        Remove-Item $dirs[0].FullName -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        foreach ($d in $dirs) { Move-ContentsSilent $d.FullName $DestRoot }
        foreach ($f in $files) {
            try { Move-Item $f.FullName $DestRoot -Force -ErrorAction Stop } catch { }
        }
        foreach ($d in $dirs) {
            Remove-Item $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Unblock files (silent)
try {
    Get-ChildItem $DestRoot -Recurse -File -Force |
        ForEach-Object { try { Unblock-File $_.FullName -ErrorAction SilentlyContinue } catch { } }
} catch { }

# Launch toolbox (same session)
if (Test-Path $Launcher) {
    . $Launcher
}
