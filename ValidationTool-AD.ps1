# ╔═════════════════════════════════════════════════════════════╗
# ║ 🧰 CS Tech Toolbox – AD Validation Tool                    ║
# ║ Version: AD.2 | 2025-07-21                                  ║
# ║ Includes: AD Users, Groups, Computers, OUs, ZIP + Cleanup   ║
# ╚═════════════════════════════════════════════════════════════╝

irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
Ensure-ExportFolder

function Run-ADUsers {
    Show-Header "Active Directory Users"
    if (-not (Get-Command Get-ADUser -ErrorAction SilentlyContinue)) {
        Write-Host "Active Directory module not available." -ForegroundColor Red
        Pause-Script; return
    }

    try {
        $users = Get-ADUser -Filter * -Properties DisplayName, EmailAddress, Enabled, LastLogonDate, Department |
            Select-Object SamAccountName, DisplayName, EmailAddress, Enabled, LastLogonDate, Department
        Export-Data -Object $users -BaseName "AD_Users"
    } catch {
        Write-Host "Failed to query AD users." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADGroups {
    Show-Header "Active Directory Groups"
    if (-not (Get-Command Get-ADGroup -ErrorAction SilentlyContinue)) {
        Write-Host "Active Directory module not available." -ForegroundColor Red
        Pause-Script; return
    }

    try {
        $groups = Get-ADGroup -Filter * -Properties Description, GroupScope |
            Select-Object Name, GroupScope, Description
        Export-Data -Object $groups -BaseName "AD_Groups"
    } catch {
        Write-Host "Failed to query AD groups." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADComputers {
    Show-Header "Active Directory Computers"
    if (-not (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
        Write-Host "Active Directory module not available." -ForegroundColor Red
        Pause-Script; return
    }

    try {
        $computers = Get-ADComputer -Filter * -Properties IPv4Address, OperatingSystem, LastLogonDate |
            Select-Object Name, IPv4Address, OperatingSystem, LastLogonDate
        Export-Data -Object $computers -BaseName "AD_Computers"
    } catch {
        Write-Host "Failed to query AD computers." -ForegroundColor Red
    }
    Pause-Script
}

function Run-ADOUs {
    Show-Header "Active Directory Organizational Units"
    if (-not (Get-Command Get-ADOrganizationalUnit -ErrorAction SilentlyContinue)) {
        Write-Host "Active Directory module not available." -ForegroundColor Red
        Pause-Script; return
    }

    try {
        $ous = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName |
            Select-Object Name, DistinguishedName, CanonicalName
        Export-Data -Object $ous -BaseName "AD_OUs"
    } catch {
        Write-Host "Failed to retrieve OU data." -ForegroundColor Red
    }
    Pause-Script
}

function Show-ADMenu {
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   🧰 CS Tech Toolbox – AD Validation Tool    ║" -ForegroundColor Cyan
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
        "5" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-ZipAndEmailResults
        }
        "6" {
            irm https://raw.githubusercontent.com/dmooney-cs/prod/refs/heads/main/Functions-Common.ps1 | iex
            Invoke-CleanupExportFolder
        }
        "Q" { return }
        default {
            Write-Host "Invalid selection." -ForegroundColor Red
            Pause-Script
        }
    }
    Show-ADMenu
}

Show-ADMenu
