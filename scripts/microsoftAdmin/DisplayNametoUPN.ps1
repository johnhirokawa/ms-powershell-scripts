# This script retrieves User Principal Names (UPNs) from Microsoft Graph based on DisplayNames listed in a CSV and exports the results to a new CSV.

# Ensure Microsoft Graph module is installed
#Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

# Define input and output file paths
$inputCsv = "\\...\ServiceAccounts_DisplayNames.csv"
$outputCsv = "\\...\ServiceAccounts_UPN.csv"

# Read input CSV
$users = Import-Csv -Path $inputCsv

# Prepare results array
$results = @()

# Loop through each DisplayName and get UPN
foreach ($user in $users) {
    $displayName = $user.DisplayName
    $mgUser = Get-MgUser -Filter "displayName eq '$displayName'" -Property UserPrincipalName,DisplayName
    if ($mgUser) {
        $results += [PSCustomObject]@{
            DisplayName = $mgUser.DisplayName
            UPN         = $mgUser.UserPrincipalName
        }
    } else {
        $results += [PSCustomObject]@{
            DisplayName = $displayName
            UPN         = "Not Found"
        }
    }
}

# Export results to CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "UPNs exported to $outputCsv"
