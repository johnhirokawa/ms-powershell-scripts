# This script connects to Exchange Online using an Entra-registered PowerShell App and performs mailbox queries.

function Connect-ToExchange {
    # Import necessary modules
    Import-Module ExchangeOnlineManagement
    Import-Module CredentialManager
    Import-Module ImportExcel

    # Set execution policy
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

    # Retrieve stored credentials for Exchange Online API
    $global:ExchangeAPICred = Get-StoredCredential -Target "ExchangeAPI"

    # Define variables
    $global:TenantId = "[TenantID]"
    $global:ClientId = $ExchangeAPICred.UserName

    # Convert SecureString to plain text for Client Secret
    $global:ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ExchangeAPICred.Password)
    )

    # Debug output
    Write-Host "Client ID: $ClientId"

    # Construct the body for the token request
    $Body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://outlook.office365.com/.default"
        Client_Id     = $ClientId
        Client_Secret = $ClientSecret
    }

    # Get the token from Microsoft Identity platform
    $global:Token = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -ContentType "application/x-www-form-urlencoded" -Body $Body

    # Check if the token was retrieved successfully
    if ($null -ne $Token) {
        Write-Host "Access token retrieved successfully."

        # Use the access token to connect to Exchange Online
        Connect-ExchangeOnline -AccessToken $Token.access_token -Organization $TenantId -ShowProgress $false
    } else {
        Write-Host "Error retrieving token."
    }
}

function ExchangeUpdates {
    Get-Mailbox -Identity "[YourUPN]"
}

# Call the function
Connect-ToExchange
ExchangeUpdates

# Access the global variables outside the function
Write-Host "Tenant ID: $TenantId"
Write-Host "Client ID: $ClientId"

Disconnect-ExchangeOnline -Confirm:$false
