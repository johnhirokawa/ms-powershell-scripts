# This script creates a distribution list in Exchange Online and adds members based on first and last names from an Excel file.

# ============================
# First-Time Setup (Uncomment if needed)
# ============================
# Install-Module -Name ImportExcel
# Install-Module -Name ExchangeOnlineManagement

# ============================
# Configuration Section
# ============================
$ExcelFilePath = "[PathToExcelFile]" # <-- Change this to the Excel file containing membership info
$SheetName = "[WorksheetName]"
$DLName = "[DistributionListName]" 
$DLDisplayName = "[DistributionListDisplayName]"
$DLAlias = "[DistributionListAlias]"
$DLPrimarySMTP = "[DistributionListPrimarySMTP]"
$Domain = "[EmailDomain]"

# Set the DL Owner (can be UPN, alias, or display name)
$DLOwner = "[DistributionListOwnerUPN]"

# ============================
# Script Logic
# ============================

# Import Excel & Exchange modules / Connect to Exchange
Import-Module ImportExcel -ErrorAction SilentlyContinue
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName "[YourUPN]"

# Read Excel data
$UserList = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName

# Create the distribution list if it doesn't exist
if (-not (Get-DistributionGroup -Identity $DLName -ErrorAction SilentlyContinue)) {
    New-DistributionGroup -Name $DLName -DisplayName $DLDisplayName -Alias $DLAlias -PrimarySmtpAddress $DLPrimarySMTP -ManagedBy $DLOwner
    Write-Host "Distribution list '$DLName' created and managed by $DLOwner"
} else {
    Write-Host "Distribution list '$DLName' already exists."
}

# Add members to the distribution list
foreach ($user in $UserList) {
    $firstName = $user.FirstName
    $lastName = $user.LastName

    # Search for user in directory
    $matchedUser = Get-Recipient -Filter "FirstName -eq '$firstName' -and LastName -eq '$lastName'" -ErrorAction SilentlyContinue
    if ($matchedUser) {
        try {
            Add-DistributionGroupMember -Identity $DLName -Member $matchedUser.PrimarySmtpAddress -ErrorAction Stop
            Write-Host "Added $($matchedUser.PrimarySmtpAddress) to $DLName"
        } catch {
            Write-Warning "Failed to add $($matchedUser.PrimarySmtpAddress): $_"
        }
    } else {
        Write-Warning "User not found: $firstName $lastName"
    }
}
