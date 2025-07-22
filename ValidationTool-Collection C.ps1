# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Validation Tool C                      ║
# ║ Version: C.2 | 2025-07-21                                   ║
# ║ Includes OSQuery + TLS Cipher + ZIP + Cleanup              ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-OSQueryBrowserExtensions {
    Show-Header "OSQuery Browser Extensions Audit"
    $paths = @("C:\Program Files (x86)\CyberCNSAgent\osqueryi.exe", "C:\Windows\CyberCNSAgent\osqueryi.exe")
    $osquery = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $osquery) {
        Write-Host "OSQuery not found." -ForegroundColor Red
        Pause-Script; return
    }

    $query = @"
SELECT
  users.username AS user,
  chrome_extensions.name AS extension_name,
  chrome_extensions.version AS extension_version,
  chrome_extensions.description AS extension_description,
  chrome_extensions.identifier AS extension_id
FROM
  chrome_extensions
JOIN users ON chrome_extensions.uid = users.uid
ORDER BY users.username, chrome_extensions.name;
"@

    try {
        $raw = & "$osquery" --json "$query"
        $parsed = $raw | ConvertFrom-Json
        Export-Data -Object $parsed -BaseName "OSQueryBrowserExtensions" -Ext "json"
        Export-Data -Object $parsed -BaseName "OSQueryBrowserExtensions" -Ext "csv"
    } catch {
        Write-Host "OSQuery execution failed." -ForegroundColor Red
    }
    Pause-Script
}

function Run-SSLCipherValidation {
    Show-Header "SSL Cipher Validation (Nmap Port 443)"
    $nmap = "C:\Program Files (x86)\CyberCNSAgent\nmap\nmap.exe"
    if (-not (Test-Path $nmap)) {
        Write-Host "Nmap not found at: $nmap" -ForegroundColor Red
        Pause-Script; return
    }

    $ip = Read-Host "Enter target IP for scan"
    $log = Get-ExportPath -BaseName "SSLCipher443Scan-$ip" -Ext "txt"

    try {
        $result = & $nmap --script ssl-enum-ciphers -p 443 $ip
        $result | Out-File $log -Encoding UTF8
        Write-ExportPath $log
    } catch {
        Write-Host "Scan failed." -ForegroundColor Red
    }

    Pause-Script
}

function Show-CollectionMenuC {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – Collection Tool C     ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] OSQuery Browser Extensions",
        " [2] SSL Cipher Validation (443)",
        " [3] Zip and Email Results",
        " [4] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-OSQueryBrowserExtensions }
        "2" { Run-SSLCipherValidation }
        "3" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-ZipAndEmailResults
        }
        "4" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-CleanupExportFolder
        }
        "Q" { return }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Pause-Script
        }
    }
    Show-CollectionMenuC
}

Show-CollectionMenuC
