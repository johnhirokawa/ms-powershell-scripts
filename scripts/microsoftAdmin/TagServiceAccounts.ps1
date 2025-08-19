# Ensure Microsoft Graph module is installed
#Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber

# Connect to Microsoft Graph with required permissions
#Connect-MgGraph -Scopes "User.ReadWrite.All"

#Define CSV path for csv export of all entra service accounts
$csvPath = "folder\enterCSVfilepath.csv"

# Import the CSV
$users = Import-Csv -Path $csvPath

# Loop through each user and update extensionAttribute1
foreach ($user in $users) {
    $upn = $user.UPN
    $displayName = $user.DisplayName

    try {
        # Find the user by UPN
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties extensionAttribute1

        if ($adUser) {
            # Update extensionAttribute1
            Set-ADUser -Identity $adUser.DistinguishedName -Add @{extensionAttribute1 = "ServiceAccount"}
            Write-Host "Updated $displayName ($upn) successfully." -ForegroundColor Green
        } else {
            Write-Host "User not found: $displayName ($upn)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Failed to update $displayName ($upn): $_" -ForegroundColor Red
    }
}
