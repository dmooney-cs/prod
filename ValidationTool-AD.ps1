
# ╔════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Active Directory Tool           ║
# ╚════════════════════════════════════════════════════╝

function Collect-ADUsers {
    Write-Host "▶ Collecting Active Directory Users..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\AD_Users_$ts_$hn.csv"
    $users = Get-ADUser -Filter * -Properties * | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet, Department, Title, EmailAddress
    $users | Export-Csv -Path $out -NoTypeInformation
    Write-Host "✔ Users exported to: $out" -ForegroundColor Green
}

function Collect-ADGroups {
    Write-Host "▶ Collecting Active Directory Groups..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\AD_Groups_$ts_$hn.csv"
    $groups = Get-ADGroup -Filter * -Properties * | Select-Object Name, GroupCategory, GroupScope, Description, WhenCreated
    $groups | Export-Csv -Path $out -NoTypeInformation
    Write-Host "✔ Groups exported to: $out" -ForegroundColor Green
}

function Collect-ADOrganizationalUnits {
    Write-Host "▶ Collecting Active Directory OUs..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\AD_OUs_$ts_$hn.csv"
    $ous = Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName, ProtectedFromAccidentalDeletion, WhenCreated
    $ous | Export-Csv -Path $out -NoTypeInformation
    Write-Host "✔ OUs exported to: $out" -ForegroundColor Green
}

function Collect-ADGPOs {
    Write-Host "▶ Collecting Active Directory GPOs..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\AD_GPOs_$ts_$hn.csv"
    $gpos = Get-GPO -All | Select-Object DisplayName, Id, Owner, CreationTime, ModificationTime, GpoStatus
    $gpos | Export-Csv -Path $out -NoTypeInformation
    Write-Host "✔ GPOs exported to: $out" -ForegroundColor Green
}

function Collect-ADComputers {
    Write-Host "▶ Collecting Active Directory Computers..." -ForegroundColor Cyan
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $hn = $env:COMPUTERNAME
    $out = "C:\Script-Export\AD_Computers_$ts_$hn.csv"
    $comps = Get-ADComputer -Filter * -Properties * | Select-Object Name, DNSHostName, OperatingSystem, Enabled, LastLogonDate, Created
    $comps | Export-Csv -Path $out -NoTypeInformation
    Write-Host "✔ Computers exported to: $out" -ForegroundColor Green
}

function Run-AllADCollections {
    Collect-ADUsers
    Collect-ADGroups
    Collect-ADOrganizationalUnits
    Collect-ADGPOs
    Collect-ADComputers
    Write-Host "✔ All AD data collected." -ForegroundColor Green
}

function Run-ZipAndEmailResults {
    $exportDir = "C:\Script-Export"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $hostname = $env:COMPUTERNAME
    $zipFile = "$exportDir\ADExport_$hostname_$timestamp.zip"

    if (Test-Path $exportDir) {
        Compress-Archive -Path "$exportDir\*" -DestinationPath $zipFile -Force
        Write-Host "✔ ZIP file created: $zipFile" -ForegroundColor Green

        $to = Read-Host "Enter recipient email"
        $mailto = "mailto:$to?subject=AD%20Export%20Results&body=Results%20attached.%20ZIP:%20$zipFile"
        Start-Process $mailto
    } else {
        Write-Host "❌ No export folder found." -ForegroundColor Yellow
    }
}

function Run-CleanupScriptData {
    Write-Host "🧹 Cleaning up export folder..." -ForegroundColor Red
    Remove-Item -Path "C:\Script-Export\*" -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "✔ Cleanup complete."
}

function Show-ADMenu {
    do {
        Write-Host "`n========= 👥 Active Directory Tool Menu ==========" -ForegroundColor Cyan
        Write-Host "[1] Collect AD Users"
        Write-Host "[2] Collect AD Groups"
        Write-Host "[3] Collect Organizational Units"
        Write-Host "[4] Collect GPOs"
        Write-Host "[5] Collect AD Computers"
        Write-Host "[6] Run All Collections"
        Write-Host "[7] Zip and Email Results"
        Write-Host "[8] Cleanup Export Folder"
        Write-Host "[Q] Quit"

        $choice = Read-Host "Select an option"
        switch ($choice) {
            "1" { Collect-ADUsers }
            "2" { Collect-ADGroups }
            "3" { Collect-ADOrganizationalUnits }
            "4" { Collect-ADGPOs }
            "5" { Collect-ADComputers }
            "6" { Run-AllADCollections }
            "7" { Run-ZipAndEmailResults }
            "8" { Run-CleanupScriptData }
            "Q" { return }
            default { Write-Host "Invalid option. Try again." -ForegroundColor Yellow }
        }
    } while ($true)
}

Show-ADMenu
