# This script identifies users with sign-in blocked and Office 365 E3 licenses, then optionally downgrades them to Microsoft 365 Business Basic.

# Connect to Microsoft 365
Connect-MsolService

# Get users with sign-in blocked and Office 365 E3 license
$blockedUsers = Get-MsolUser -All | Where-Object {
    $_.BlockCredential -eq $true -and $_.Licenses.AccountSkuId -contains "[TenantPrefix]:ENTERPRISEPACK"
}

# Display the list of users
$blockedUsers | Select-Object UserPrincipalName, DisplayName

# Confirm action
$confirmation = Read-Host "Do you want to move these users to Microsoft 365 Business Basic? (y/n)"

if ($confirmation -eq 'y') {
    foreach ($user in $blockedUsers) {
        # Remove Office 365 E3 license
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses "[TenantPrefix]:ENTERPRISEPACK"

        # Add Microsoft 365 Business Basic license
        Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses "[TenantPrefix]:BUSINESS_BASIC"
    }
    Write-Host "Users have been moved to Microsoft 365 Business Basic."
} else {
    Write-Host "Operation cancelled."
}

# Optional: Testing query
Get-MsolUser -All | Where-Object {
    $_.BlockCredential -eq $true -and $_.Licenses.AccountSkuId -contains "ENTERPRISEPACK"
} | Select-Object UserPrincipalName, DisplayName
