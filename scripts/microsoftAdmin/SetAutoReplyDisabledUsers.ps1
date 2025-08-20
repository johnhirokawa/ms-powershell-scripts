# This script sets auto-reply messages in Exchange Online for disabled users across specified OUs.

# Import required modules
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName [YourUPN] -ShowProgress $true

# Define the auto-reply message
$autoReplyMessage = @"
Aloha,<br><br> Mahalo for reaching out to [Insert Company Name]!<br><br> 
Please note that the employee you are trying to contact is <b>no longer with our organization.</b><br><br> 
For immediate assistance, please contact our main line at [Insert Phone Number] or email us at [InfoEmail].<br> 
Our Admin team will be happy to direct your inquiry to the appropriate department.<br><br> 
Mahalo,<br> [Insert Company Name]
"@

# Define the list of specific OUs
$OUs = @(
    "OU=AD OU 1,OU=OU,DC=[LocalDomain],DC=local",
    "OU=AD OU 2,OU=OU,DC=[LocalDomain],DC=local",
)

# Collect disabled users
$disabledUsers = @()
foreach ($OU in $OUs) {
    $disabledUsers += Get-ADUser -Filter { Enabled -eq $false -and EmailAddress -like "*" } -SearchBase $OU -Properties EmailAddress, UserPrincipalName
}

# Display disabled users
$disabledUsers | Select-Object -Property Name, SamAccountName

# Initialize tracking lists
$successList = @()
$failureList = @()

# Set auto-reply for each disabled user
foreach ($user in $disabledUsers) {
    $userPrincipalName = $user.UserPrincipalName
    if ($userPrincipalName) {
        try {
            Set-MailboxAutoReplyConfiguration -Identity $userPrincipalName -AutoReplyState Enabled -InternalMessage $autoReplyMessage -ExternalMessage $autoReplyMessage
            $successList += $user.Name
        } catch {
            $failureList += [PSCustomObject]@{
                Name   = $user.Name
                Reason = $_.Exception.Message
            }
        }
    } else {
        $failureList += [PSCustomObject]@{
            Name   = $user.Name
            Reason = "UserPrincipalName not found"
        }
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

# Output results
Write-Output "Successfully changed users:"
$successList

Write-Output "Failed to change users and reasons:"
$failureList | Format-Table -AutoSize

# Optional: Script to get all child OUs under a parent OU
<# 
$parentOU = "OU=OU,DC=[LocalDomain],DC=local"
$childOUs = Get-ADOrganizationalUnit -SearchBase $parentOU -Filter *
$formattedOUs = $childOUs | ForEach-Object { '"{0}"' -f $_.DistinguishedName }
$formattedOUsString = $formattedOUs -join ",`n"
Write-Output $formattedOUsString
#>
