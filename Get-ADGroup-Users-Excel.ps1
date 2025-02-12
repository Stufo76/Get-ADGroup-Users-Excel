<#
.SYNOPSIS
    This PowerShell script retrieves a list of users from an Active Directory group
    and exports the results to an Excel file.

.DESCRIPTION
    - Prompts the user to enter an Active Directory group name.
    - Checks if the group exists in Active Directory.
    - Retrieves all users from the group, including their first name, last name, and account status (Active/Disabled).
    - Displays the result in a formatted table in the PowerShell console.
    - Automatically saves the output to an Excel file (.xlsx) with auto-sizing, bold headers, a frozen top row, and an auto-filter.

.AUTHOR
    Diego Pastore "Stufo76" (stufo76@gmail.com)

.LICENSE
    This script is released under the GNU General Public License v3.0.
    You are free to use, modify, and distribute it under the terms of the GPL-3.0 license.

.REQUIREMENTS
    - Windows PowerShell 5.1 or later.
    - Active Directory PowerShell module.
    - ImportExcel PowerShell module (installed automatically if missing).

.NOTES
    - The Excel file is saved to "C:\Temp\Users_<GroupName>.xlsx".
    - The script requires administrator privileges to query Active Directory.
    - The ImportExcel module is used to generate the Excel file without requiring Microsoft Excel to be installed.

.VERSION
    1.0 - Initial release.

#>

# Import the Excel module (install it if not available)
$ExcelModule = Get-Module -ListAvailable -Name ImportExcel
if (-not $ExcelModule) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Prompt the user to enter the Active Directory group name
$GroupName = Read-Host "Enter the Active Directory group name"

# Check if the group exists
$Group = Get-ADGroup -Filter { Name -eq $GroupName } -ErrorAction SilentlyContinue

if (-not $Group) {
    Write-Host "Error: The group '$GroupName' does not exist in Active Directory." -ForegroundColor Red
    exit
}

# Retrieve group members
$Members = Get-ADGroupMember -Identity $GroupName -Recursive | Where-Object { $_.objectClass -eq "user" }

# Check if the group has any users
if (-not $Members) {
    Write-Host "The group '$GroupName' has no users." -ForegroundColor Yellow
    exit
}

# Extract user details including status (Active/Disabled)
$UserList = $Members | ForEach-Object {
    $User = Get-ADUser -Identity $_.DistinguishedName -Properties GivenName, Surname, Enabled
    [PSCustomObject]@{
        LastName  = $User.Surname
        FirstName = $User.GivenName
        Status    = if ($User.Enabled) { "Active" } else { "Disabled" }
    }
}

# Sort the user list
$UserList = $UserList | Sort-Object LastName, FirstName

# Format output as a table string for easy copy-paste
$FormattedTable = $UserList | Format-Table -AutoSize | Out-String

# Display the formatted table
Write-Host "`n--- USER LIST IN GROUP: $GroupName ---" -ForegroundColor Cyan
Write-Host $FormattedTable

# Define Excel file path
$ExcelPath = "C:\Temp\Users_$GroupName.xlsx"

# Save the output to an Excel file
$UserList | Export-Excel -Path $ExcelPath -WorksheetName "Users" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter

Write-Host "Excel file saved at: $ExcelPath" -ForegroundColor Green
