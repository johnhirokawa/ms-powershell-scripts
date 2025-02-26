### Exchange - Add Bulk Emails to Distribution List ###

#Run Install-Module Command on first run for computer (Uncomment it out to run on 1st run)
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline -UserPrincipalName [Enter Admin Login Email]

# Define the distribution list and the emails to add
$DistributionList = "Distribution_List_Email@Domain.com"


$EmailsToAdd = @(
    "email_1@domain.com",
    "email_2@domain.com",
    "email_3@domain.com",
    "email_4@domain.com"
)

# Add each email to the distribution list
foreach ($Email in $EmailsToAdd) {
    Add-DistributionGroupMember -Identity $DistributionList -Member $Email
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

### Code Credit: John Hirokawa 
