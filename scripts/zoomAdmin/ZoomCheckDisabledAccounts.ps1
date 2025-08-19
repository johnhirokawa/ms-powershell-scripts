# This script checks disabled AD users across multiple OUs and compares their Zoom account status, summarizing discrepancies.

# Install required modules (uncomment if needed)
# Install-Module CredentialManager
# Install-Module PSZoom

Import-Module CredentialManager
Import-Module PSZoom

# Define the OU map
$ouMap = @{
    "AD OU 1"       = "OU=AD OU 1,DC=[LocalDomain],DC=local"
    "AD OU 2"          = "OU=AD OU 2,DC=[LocalDomain],DC=local"
}

function Get-DisabledUsersFromOU {
    param ([string]$ou)
    Get-ADUser -Filter {Enabled -eq $false} -SearchBase $ou -Properties DisplayName, Enabled, SamAccountName |
        Select-Object DisplayName, Enabled, SamAccountName
}

function Check-ZoomUserStatus {
    param ([string]$email)
    try {
        $zoomUser = Get-ZoomUser -Email $email -ErrorAction Stop
        return $zoomUser.Status -eq 'inactive'
    } catch {
        if ($_.Exception -match "404") {
            Write-Output "User with email $email not found in Zoom."
            return $null
        } else {
            throw $_
        }
    }
}

# Connect to Zoom API
$ZoomAPI = Get-StoredCredential -Target "ZoomAPI"
$zclientId = $ZoomAPI.UserName
$zclientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ZoomAPI.Password)
)
Connect-PSZoom -AccountID '[ZoomAccountId]' -ClientID $zclientId -ClientSecret $zclientSecret

# Initialize counters
$notFoundInZoomCount = 0
$deactivatedInBothCount = 0
$deactivatedInADActiveInZoomCount = 0

# Iterate through each OU and check disabled users in Zoom
foreach ($ou in $ouMap.Values) {
    $disabledUsers = Get-DisabledUsersFromOU -ou $ou
    foreach ($user in $disabledUsers) {
        $email = "$($user.SamAccountName)@[EmailDomain]"
        $isZoomDeactivated = Check-ZoomUserStatus -email $email

        if ($isZoomDeactivated -eq $true) {
            Write-Output "User $($user.DisplayName) is deactivated in both AD and Zoom."
            $deactivatedInBothCount++
        } elseif ($isZoomDeactivated -eq $false) {
            Write-Output "User $($user.DisplayName) is deactivated in AD but active in Zoom."
            $deactivatedInADActiveInZoomCount++
        } elseif ($isZoomDeactivated -eq $null) {
            $notFoundInZoomCount++
        }
    }
    Write-Output ""
}

# Output the summary
Write-Output "Summary:"
Write-Output "Users not found in Zoom: $notFoundInZoomCount"
Write-Output "Users deactivated in both AD and Zoom: $deactivatedInBothCount"
Write-Output "Users deactivated in AD but active in Zoom: $deactivatedInADActiveInZoomCount"
