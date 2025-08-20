# This script assigns a specific Microsoft 365 license to users listed in a CSV file using Microsoft Graph.

# Define tenant and license information
$TenantName = "[TenantDomain]"
$LicenseSKU = "[LicenseSkuId]"
$TenantLicense = "${TenantName}:$LicenseSKU"
$TenantLicense

# Define the path to the CSV file
$FilePath = "[PathToCsvFile]"

# Connect to Microsoft Graph
Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

# Import the list of users from the CSV file
$Users = Import-Csv -Path $FilePath

# Assign the license to each user
foreach ($User in $Users) {
    Set-MgUserLicense -UserId $User.UserPrincipalName -AddLicenses @{SkuId = $LicenseSKU} -RemoveLicenses @()
}

# Output the result
Write-Output "Licenses have been assigned."
