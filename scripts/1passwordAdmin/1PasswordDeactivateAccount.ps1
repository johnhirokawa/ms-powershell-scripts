<#
Serice Account Token: 
#>

# Define variables
$Env:OP_SERVICE_ACCOUNT_TOKEN = "[Insert API Token here]"
$userToDeactivate = "[Email of User to Deactivate]"

# Authenticate using the service account token
$sessionToken = & op signin --raw --session=$serviceAccountToken

# Deactivate the user
& 
op user suspend $userToDeactivate
# Clear the session token
$sessionToken = $null

Write-Output "User $userToDeactivate has been deactivated."
