# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – Active Directory Validation            ║
# ║ Version: AD.1 | 2025-07-21                                  ║
# ║ Includes: AD Users, Groups, Computers, OUs, ZIP + Cleanup   ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-ADUsers {
    Show-Header "Active Directory Users"
    try {
        $users = Get-ADUser -Filter * -Properties DisplayName, EmailAddress, Enabled, LastLogonDate, Department |
            Select-Object SamAccountName, DisplayName, EmailAddress, Enabled, LastLogonDate, Department
        Export-Data -Object $users -BaseName "AD_Users"
    } catch {
        Write-Host "AD module not available or access denied." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADGroups {
    Show-Header "Active Directory Groups"
    try {
        $groups = Get-ADGroup -Filter * -Properties Description, GroupScope, Members |
            Select-Object Name, GroupScope, Description
        Export-Data -Object $groups -BaseName "AD_Groups"
    } catch {
        Write-Host "Failed to query AD groups." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADComputers {
    Show-Header "Active Directory Computers"
    try {
        $computers = Get-ADComputer -Filter * -Properties IPv4Address, OperatingSystem, LastLogonDate |
            Select-Object Name, IPv4Address, OperatingSystem, LastLogonDate
        Export-Data -Object $computers -BaseName "AD_Computers"
    } catch {
        Write-Host "Unable to query computer objects." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADOUs {
    Show-Header "Active Directory Organizational Units"
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName, Created, Modified |
            Select-Object Name, DistinguishedName, CanonicalName
        Export-Data -Object $ous -BaseName "AD_OUs"
    } catch {
        Write-Host "Unable to retrieve OU data." -ForegroundColor Red
    }
    Pause-Script
}

function Show-ADMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – AD Collection Tool    ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    $menu = @(
        " [1] Collect AD Users",
        " [2] Collect AD Groups",
        " [3] Collect AD Computers",
        " [4] Collect AD Organizational Units",
        " [5] Zip and Email Results",
        " [6] Cleanup Export Folder",
        " [Q] Quit"
    )
    $menu | ForEach-Object { Write-Host $_ }

    $sel = Read-Host "`nSelect an option"
    switch ($sel) {
        "1" { Run-ADUsers }
        "2" { Run-ADGroups }
        "3" { Run-ADComputers }
        "4" { Run-ADOUs }
        "5" { Invoke-ZipAndEmailResults }
        "6" { Invoke-CleanupExportFolder }
        "Q" { return }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause-Script }
    }
    Show-ADMenu
}

Show-ADMenu
