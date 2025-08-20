# This script deactivates a Zoom user account using Zoom's REST API and credentials stored in Windows Credential Manager.

# Install required modules (uncomment if needed)
# Install-Module CredentialManager
# Install-Module PSZoom

Import-Module CredentialManager -Verbose

# Retrieve Zoom credentials from Credential Manager
$ZoomAPI = Get-StoredCredential -Target "ZoomAPI"
$ZclientId = $ZoomAPI.UserName
$ZclientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ZoomAPI.Password)
)
$ZaccountId = "[ZoomAccountId]"

# Get an access token
$ztokenUri = "https://zoom.us/oauth/token"
$ztokenHeaders = @{
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${ZclientId}:${ZclientSecret}"))
    "Content-Type"  = "application/x-www-form-urlencoded"
}
$ztokenBody = "grant_type=account_credentials&account_id=${ZaccountId}"
$ztokenResponse = Invoke-RestMethod -Uri $ztokenUri -Method Post -Headers $ztokenHeaders -Body $ztokenBody
$zaccessToken = $ztokenResponse.access_token

# Clear sensitive variables
$ZclientId = ""
$ZclientSecret = ""

# Define the Zoom user to deactivate
$zuserId = "testuser@[EmailDomain]"

# Define the API endpoint and headers
$zuri = "https://api.zoom.us/v2/users/$zuserId/status"
$zheaders = @{
    "Authorization" = "Bearer $zaccessToken"
    "Content-Type"  = "application/json"
}

# Define the request body
$zbody = @{
    action = "deactivate"
} | ConvertTo-Json

# Send the request to deactivate the user
$zresponse = Invoke-RestMethod -Uri $zuri -Method Put -Headers $zheaders -Body $zbody

# Output the response
$zresponse
