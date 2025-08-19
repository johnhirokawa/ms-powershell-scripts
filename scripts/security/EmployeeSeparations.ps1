# This script automates employee separation tasks including disabling accounts, hiding from GAL, blocking sign-in, resetting passwords, updating licenses, setting auto-replies, managing mail forwarding, and sending summary reports via email.

Write-Output "Starting Employee Separation Script..."
##Idea: What about adding a piece to disable user in mosyle?##
## Description: Powershell Script - Email Fwd & Auto Reply ##
## Client ID: [ClientID] ##
## Tenant ID: [TenantID] ##

###LOGGING - Setting up Logging
# Get the current date and time
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create the log file name with the timestamp
$logFile = "\\...\logfile_$timestamp.txt"

# Initialize the log file
New-Item -Path $logFile -ItemType File -Force

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
}

### INITIAL SETUP
#Unblock-File -Path "\\...\EmployeeSeparations.ps1" #---> This is used for unblocking the file.

Log-Message "PowerShell version: $($PSVersionTable.PSVersion)"
Log-Message "Execution Policy: $(Get-ExecutionPolicy)"

Log-Message "Starting Employee Separation Script..."
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

##NOTE TO SELF - Install-Module Microsoft.Graph takes really long and should only be run once. Also want to make sure and not run Import-Module

# Import necessary modules
<#Install-Module ActiveDirectory -Force
Install-Module ExchangeOnlineManagement -Force
Install-Module -Name PnP.PowerShell -Scope CurrentUser
Install-Module CredentialManager -Force#>
#Install-Module -Name Microsoft.Graph -Force #This install takes really long and should only be run once on the computer. Also want to make sure and not run MgGraph Import-Modules that aren't necessary for the commands being used.
Log-Message "Reached checkpoint 1..."


### Importing Modules - This step allows you to run the different commands. ###
Log-Message "Importing necessary powershell modules...."
function Module_Setup{
    try{
        Import-Module ActiveDirectory -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module ExchangeOnlineManagement -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module CredentialManager -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module ImportExcel
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module Microsoft.Graph.Authentication -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module Microsoft.Graph.Users -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
    try{
        Import-Module Microsoft.Graph.Files -Verbose
    } catch{
        Log-Message "An error occurred: $_"
    }
}

Module_Setup

Log-Message "PowerShell Module import process completed..."

Log-Message "Reached checkpoint 2..."

#This function authenticates with MS Graph API for any commands that utilize MS Graph
function Mg-GraphConnect {
    
    # Retrieve stored credentials
    $global:ExchangeAPICred = Get-StoredCredential -Target "ExchangeAPI"

    # Check if credentials are retrieved
    if ($null -eq $global:ExchangeAPICred) {
        Write-Host "Failed to retrieve credentials. Please check the target name."
        return
    }

    # Define variables
    $global:TenantId = "[TenantID]"
    $global:ClientId = $global:ExchangeAPICred.UserName

    # Convert SecureString to plain text for Client Secret
    $plainClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:ExchangeAPICred.Password))
    $global:ClientSecret = ConvertTo-SecureString -String $plainClientSecret -AsPlainText -Force

    # Debug output
    Write-Host "Client ID: $global:ClientId"

    # Check if ClientId and ClientSecret are not null or empty
    if ([string]::IsNullOrEmpty($global:ClientId) -or [string]::IsNullOrEmpty($plainClientSecret)) {
        Write-Host "Client ID or Client Secret is null or empty. Please check your credentials."
        return
    }

    # Get an access token
    $token = Get-MsalToken -TenantId $global:TenantId -ClientId $global:ClientId -ClientSecret $global:ClientSecret -Scopes "https://graph.microsoft.com/.default"

    # Check if token is retrieved
    if ($null -eq $token) {
        Write-Host "Failed to retrieve access token."
        return
    }

    # Convert the access token to a SecureString
    $global:AccessToken = $token.AccessToken 

    # Connect to Microsoft Graph
    Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $global:AccessToken -AsPlainText -Force)
    Write-Host "Connection Completed"
}

#This function authenticates with exchange for any commands that utilize Exchange
function Connect-ToExchange {
    # Import necessary modules
    Import-Module ExchangeOnlineManagement
    Import-Module CredentialManager
    Import-Module -Name ImportExcel

    # Set execution policy
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

    # Retrieve stored credentials for Exchange Online API
    $global:ExchangeAPICred = Get-StoredCredential -Target "ExchangeAPI"

    # Define variables
    $global:TenantId = "[TenantID]"
    $global:ClientId = $ExchangeAPICred.UserName

    # Convert SecureString to plain text for Client Secret
    $global:ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ExchangeAPICred.Password))

    # Debug output
    Write-Host "Client ID: $ClientId"

    # Construct the body for the token request
    $Body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://outlook.office365.com/.default"
        Client_Id     = $ClientId
        Client_Secret = $ClientSecret
    }

    try {
        # Get the token from Microsoft Identity platform
        $global:Token = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -ContentType "application/x-www-form-urlencoded" -Body $Body
        Write-Host "Token response: $($global:Token | Out-String)"

        # Check if the token was retrieved successfully
        if ($global:Token -ne $null) {
            Write-Host "Access token retrieved successfully."

            # Use the access token to connect to Exchange Online
            Connect-ExchangeOnline -AccessToken $Token.access_token -Organization $TenantId -ShowProgress $false
        } else {
            Write-Host "Error: Token is null."
        }
    } catch {
        Write-Host "An error occurred: $_"
    }
}
#This function downloads the excel file temporarily to use when script is run
<#function DownloadExcel{
    # Define the shareable link to the Excel file
    $shareableLink = "[Insert shareable download link]"  # Replace with the updated shareable link

    # Define the local path to save the file
    $localFilePath = "\\...\Employee Separations Data.xlsx"

    # Download the file
    Invoke-WebRequest -Uri $shareableLink -OutFile $localFilePath

    # Verify the file size
    $fileInfo = Get-Item $localFilePath
    Write-Host "File size: $($fileInfo.Length) bytes"

    # Check if the file is not empty
    if ($fileInfo.Length -eq 0) {
        Write-Host "The downloaded file is empty. Please check the shareable link."
        return
    }

    # Load the Excel file using ImportExcel module (if installed)
    $global:data = Import-Excel -Path $localFilePath

    # Output the data
    $global:data | ForEach-Object { $_ }
}#>

function DownloadExcel{
    # Define the user ID and file ID
    $global:OneDriveUserId = "[OneDriveUserId]"
    $global:OneDriveFileId = "[OneDriveFileId]"
    $global:WorksheetName = "Sheet1"  # Replace with your actual worksheet name

    # Define the URL to get the Excel file content
    $global:OneDriveUrl = "https://graph.microsoft.com/v1.0/users/$global:OneDriveUserId/drive/items/$global:OneDriveFileId/workbook/worksheets/$global:WorksheetName/usedRange"

    # Get the Excel file content
    $headers = @{
        Authorization = "Bearer $global:AccessToken"
    }
    $response = Invoke-RestMethod -Uri $global:OneDriveUrl -Headers $headers -Method Get

    # Convert the response to a usable format
    $global:data = @()
    $columns = $response.values[0]
    for ($i = 1; $i -lt $response.values.Count; $i++) {
        $row = $response.values[$i]
        $obj = @{}
        for ($j = 0; $j -lt $columns.Count; $j++) {
            $obj[$columns[$j]] = $row[$j]
        }
        $global:data += $obj
    }
}
"Reached checkpoint 3..."

#Running Function to Authenticate Mg-Graph
Mg-GraphConnect
Log-Message "Reached checkpoint 4..."

#Running Function to Authenticate Exchange
Connect-ToExchange
Log-Message "Reached checkpoint 5..."

#Running Function to Download current Excel document
DownloadExcel
Log-Message "Reached checkpoint 6..."

# Initialize a list to store the summary of actions
$summary = @()
Log-Message "Reached checkpoint 7..."

# Processes each row of the excel sheet and fills in the variables accordingly
foreach ($row in $global:data) {
    # Check if the rows are empty
    if (-not $row.EmployeeUsername -and -not $row.ManagerUsername) {
        Log-Message "Skipping empty row..."
        continue
    }
    $employeeName = $row.EmployeeName
    $managerName = $row.ManagerName
    $employeeUsername = $row.EmployeeUsername
    $managerUsername = $row.ManagerUsername
    $employeeUserID = $row.EmployeeUserID
    $managerUserID = $row.ManagerUserID
    $employeeEmail = $row.EmployeeEmail
    $managerEmail = $row.ManagerEmail
    $separationDateTime = [datetime]::FromOADate($row.SeparationDateTime)
    $endMailForwarding = [datetime]::FromOADate($row.EndMailForwarding)
    $emailPreference = $row.EmailPreference
    $oneDriveRequested = $row.OneDriveRequested
    $zoomDataRequested = $row.ZoomDataRequested
    $Processed_1 = $row.Processed_1
    $Processed_2 = $row.Processed_2


    $userSummary = @{
        Name = $employeeName
        SeparationDate = $separationDateTime
        ADDisabled = ""
        HiddenFromAddressBook = ""
        OfficeSignInBlocked = ""
        OfficePasswordUpdated = ""
        OfficeLicenseUpdated = ""
        EmailAutoReplySet = ""
        MailForwardingStarted = ""
        MailForwardingEnded = ""
        SharedMailboxSet = ""
        ZoomAccountDeactivated = ""
        ZoomDataTransferred = ""
    }

    # Get the current date and time at run time of script
    $currentDateTime = Get-Date

    Log-Message "Reached checkpoint 8..."
    Log-Message "**********   $employeeName   **********"
    ###If statement to say if Processed_1 is empty, then run the following code
    if(-not $Processed_1) {
        if ($currentDateTime -ge $separationDateTime) {
    ### AD - Disable user in AD if not already disabled - CONFIRMED WORKING
            try {
                $adUser = Get-ADUser -Identity $employeeUsername -Properties Enabled
                if ($adUser.Enabled -ne $false) {
                    Disable-ADAccount -Identity $employeeUsername
                    Log-Message "${employeeName}:   AD account Disabled"
                    $userSummary.ADDisabled = "X"
                } else {
                    Log-Message "${employeeName}:   AD account already Disabled"
                    $userSummary.ADDisabled = "P"
                }
            } catch {
                Log-Message "An error occurred: $_"
                $userSummary.ADDisabled = "E"
            }
        ### AD - Hide email from Global Address List if not already hidden - CONFIRMED WORKING
            try {
                $adUser = Get-ADUser -Identity $employeeUsername -Properties msExchHideFromAddressLists
                if ($adUser.msExchHideFromAddressLists -ne $true) {
                    Set-ADUser -Identity $employeeUsername -Add @{msExchHideFromAddressLists = $true}
                    Log-Message "${employeeName}:   Email hidden from Global Address List"
                    $userSummary.HiddenFromAddressBook = "X"
                } else {
                    Log-Message "${employeeName}:   Email already hidden from Global Address List"
                    $userSummary.HiddenFromAddressBook = "P"
                }
            } catch {
                Log-Message "An error occurred: $_"
                $userSummary.HiddenFromAddressBook = "E"
            }
        ### OFFICE365 - Block sign-in in Office 365 if not already blocked - CONFIRMED WORKING
            try {
                $o365User = Get-MgUser -UserId $employeeUserID -Property "AccountEnabled"

                if ($null -eq $o365User.AccountEnabled) {
                    Log-Message "${employeeName}:   AccountEnabled property is null"
                } elseif ($o365User.AccountEnabled -ne $false) {
                    Update-MgUser -UserId $employeeUserID -AccountEnabled:$false
                    Log-Message "${employeeName}:   Office 365 sign in Blocked"
                    $userSummary.OfficeSignInBlocked = "X"
        ### OFFICE365 - Change password in Office 365 to default password - CONFIRMED WORKING
                    Update-MgUser -UserId $employeeUserID -PasswordProfile @{Password="Onipaa808"; ForceChangePasswordNextSignIn=$false}
                    Log-Message "${employeeName}:   Office 365 Password has been Changed"
                    $userSummary.OfficePasswordUpdated = "X"
                }
            else {
                Log-Message "${employeeName}:   Office 365 sign-in is already Blocked"
                $userSummary.OfficeSignInBlocked = "P"
                $userSummary.OfficePasswordUpdated = "P"
                }
            } catch {
                Log-Message "An error occurred with Office Sign-in Blocked or Password Update: $_"
                $userSummary.OfficeSignInBlocked = "E"
                $userSummary.OfficePasswordUpdated = "E"
            }
        ### OFFICE365 - Downgrade License to Business Basic - Confirmed Working
            try {
                # Get the user's current licenses
                $o365UserLicenses = Get-MgUserLicenseDetail -UserId $employeeUserID

                # Define the SKU ID for Microsoft 365 Business Basic
                $businessBasicSkuId = "[BusinessBasicSkuId]"  # Replace with the actual SKU ID for Microsoft 365 Business Basic

                # Check if Microsoft 365 Business Basic is the only license
                $isBusinessBasicOnly = $true
                foreach ($license in $o365UserLicenses) {
                    if ($license.SkuId -ne $businessBasicSkuId) {
                        $isBusinessBasicOnly = $false
                        break #>This Exits the foreach loop once it sees any other license besides business basic
                    }
                }

                if (-not $isBusinessBasicOnly) {
                    # Disable all existing licenses
                    foreach ($license in $o365UserLicenses) {
                        Set-MgUserLicense -UserId $employeeUserID -RemoveLicenses @($license.SkuId) -AddLicenses @()
                        Log-Message "${employeeName}:   License ${license.SkuId} removed"
                    }

                    # Add Microsoft 365 Business Basic license
                    $assignedLicense = [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphAssignedLicense]@{SkuId = $businessBasicSkuId}
                    Set-MgUserLicense -UserId $employeeUserID -AddLicenses @($assignedLicense) -RemoveLicenses @()
                    Log-Message "${employeeName}:   Microsoft 365 Business Basic license added"
                    $userSummary.OfficeLicenseUpdated = "X"
                } else {
                    Log-Message "${employeeName}:   Microsoft 365 Business Basic is already the only license"
                    $userSummary.OfficeLicenseUpdated = "P"
                }
            } catch {
                Log-Message "An error occurred while updating licenses: $_"
                $userSummary.OfficeLicenseUpdated = "E"
            }

        ### OFFICE365 - Setting Auto-Reply - CONFIRMED WORKING
            try {
                $autoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $employeeUsername
                $internalMessage = "<html><body>Aloha,<br><br> Mahalo for reaching out to [Insert Company Name]!<br><br> Please note that the employee you are trying to contact is <b>no longer with our organization.<br><br> For immediate assistance,</b> please contact our main line at [Insert Phone Number] or email us at [Inser Email].<br> Our Admin team will be happy to direct your inquiry to the appropriate department.<br><br> Mahalo,<br> [Insert Company Name]</body></html>".Trim()
                $externalMessage = "<html><body>Aloha,<br><br> Mahalo for reaching out to [Insert Company Name]!<br><br> Please note that the employee you are trying to contact is <b>no longer with our organization.<br><br> For immediate assistance,</b> please contact our main line at [Insert Phone Number] or email us at [Insert Email.<br> Our Admin team will be happy to direct your inquiry to the appropriate department.<br><br> Mahalo,<br> [Insert Company Name]</body></html>".Trim()
            
                # Normalize whitespace for comparison
                $normalizeWhitespace = { param ($str) return -join ($str -split '\s+') }
                $currentInternalMessage = $normalizeWhitespace.Invoke($autoReplyConfig.InternalMessage)
                $currentExternalMessage = $normalizeWhitespace.Invoke($autoReplyConfig.ExternalMessage)
                $desiredInternalMessage = $normalizeWhitespace.Invoke($internalMessage)
                $desiredExternalMessage = $normalizeWhitespace.Invoke($externalMessage)
            
                # Log current and desired messages for debugging
                <#Log-Message "Current InternalMessage: $currentInternalMessage"
                Log-Message "Current ExternalMessage: $currentExternalMessage"
                Log-Message "Desired InternalMessage: $desiredInternalMessage"
                Log-Message "Desired ExternalMessage: $desiredExternalMessage"#>
            
                if ($autoReplyConfig.AutoReplyState -ne "Enabled" -or $currentInternalMessage -ne $desiredInternalMessage -or $currentExternalMessage -ne $desiredExternalMessage) {
                    Set-MailboxAutoReplyConfiguration -Identity $employeeUsername -AutoReplyState Enabled -InternalMessage $internalMessage -ExternalMessage $externalMessage
                    Log-Message "${employeeName}:   Auto-reply is Set"
                    $userSummary.EmailAutoReplySet = "X"
                } else {
                    Log-Message "${employeeName}:   Auto-reply has previously been Set"
                    $userSummary.EmailAutoReplySet = "P"
                }
            } catch {
                Log-Message "${employeeName}:   Auto-reply did NOT get Set"
                Log-Message "An error occurred: $_"
                $userSummary.EmailAutoReplySet = "E"
            }
            
        ### ZOOM - Deactivate User - Confirmed Working
            try {
                # Retrieve clientId and clientSecret from Credential Manager
                $ZoomAPI = Get-StoredCredential -Target "ZoomAPI"
            
                $zclientId = $ZoomAPI.UserName
                $zclientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ZoomAPI.Password))
                $zaccountId = "[ZoomAccountId]"
            
                # Get an access token
                $ztokenUri = "https://zoom.us/oauth/token"
                $ztokenHeaders = @{
                    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${zclientId}:${zclientSecret}"))
                    "Content-Type" = "application/x-www-form-urlencoded"
                }
                $ztokenBody = "grant_type=account_credentials&account_id=${zaccountId}"
                $ztokenResponse = Invoke-RestMethod -Uri $ztokenUri -Method Post -Headers $ztokenHeaders -Body $ztokenBody
                $zaccessToken = $ztokenResponse.access_token
                $zclientId = ""
                $zclientSecret = ""
            
                # Define the user ID of the user you want to deactivate
                $zuserId = "$employeeEmail"
            
                # Define the API endpoint and headers for checking user status
                $zstatusUri = "https://api.zoom.us/v2/users/$zuserId"
                $zstatusHeaders = @{
                    "Authorization" = "Bearer $zaccessToken"
                    "Content-Type" = "application/json"
                }
            
                # Get the current status of the user
                $zstatusResponse = Invoke-RestMethod -Uri $zstatusUri -Method Get -Headers $zstatusHeaders
            
                # Check if the user is already deactivated
                if ($zstatusResponse.status -ne "inactive") {
                    # Define the API endpoint and headers for deactivating the user
                    $zuri = "https://api.zoom.us/v2/users/$zuserId/status"
                    $zheaders = @{
                        "Authorization" = "Bearer $zaccessToken"
                        "Content-Type" = "application/json"
                    }
            
                    # Define the body of the request
                    $zbody = @{
                        action = "deactivate"
                    } | ConvertTo-Json
            
                    # Send the request to deactivate the user
                    $zresponse = Invoke-RestMethod -Uri $zuri -Method Put -Headers $zheaders -Body $zbody
            
                    # Output the response
                    $zresponse
            
                    Log-Message "${employeeName}:   Zoom User has been Deactivated"
                    $userSummary.ZoomAccountDeactivated = "X"
                } else {
                    Log-Message "${employeeName}:   Zoom User is already Deactivated"
                    $userSummary.ZoomAccountDeactivated = "P"
                }
            } catch {
                Log-Message "${employeeName}:   Zoom User has not been Deactivated"
                Log-Message "An error occurred: $_"
                $userSummary.ZoomAccountDeactivated = "E"
            }
        ###Restructure - Start
        if ($currentDateTime -lt $endMailForwarding) {
        
        ### OFFICE365 - Mail Forwarding - CONFIRMED WORKING
            if ($emailPreference -eq "Forwarding") {
                try {
                    # Get the current forwarding status
                    $mailbox = Get-Mailbox -Identity $employeeEmail
                    $currentForwardingAddress = $mailbox.ForwardingSMTPAddress
                    $currentForwardingEnabled = $mailbox.DeliverToMailboxAndForward
                    
                    # Log current forwarding status for debugging
                    <#Log-Message "Current Forwarding Address: $currentForwardingAddress"
                    Log-Message "Current Forwarding Enabled: $currentForwardingEnabled"#>
            
                # Check if forwarding is already set up correctly
                if ($currentForwardingAddress -ne "smtp:$managerEmail" -or $currentForwardingEnabled -ne $true) {
                    # Set up email forwarding
                    Set-Mailbox -Identity $employeeEmail -ForwardingSMTPAddress $managerEmail -DeliverToMailboxAndForward $true
                    Log-Message "${employeeName}:   Forwarding to Manager: $managerName - Started"
                    $userSummary.MailForwardingStarted = "X"
                } else {
                    Log-Message "${employeeName}:   Forwarding already set up to $managerName."
                    $userSummary.MailForwardingStarted = "P"
                }
                } catch {
                    Log-Message "${employeeName}:   Forwarding setup Failed"
                    Log-Message "Forwarding setup error occurred: $_"
                    $userSummary.MailForwardingStarted = "E"
                }
            } elseif ($emailPreference -eq "Shared Mailbox") {
                try {
                    # Check if employee's mailbox is already converted to Shared Mailbox
                    $mailbox = Get-Mailbox -Identity $employeeEmail
                    Log-Message "${employeeName}:   Current Mailbox Type is $($mailbox.RecipientTypeDetails)"
        
                    if ($mailbox.RecipientTypeDetails -ne "SharedMailbox") {
                        # If not, then Converting to shared mailbox
                        Set-Mailbox -Identity $employeeEmail -Type Shared
                        Log-Message "${employeeName}:   Converted Mailbox to a Shared Mailbox."
                    } else {
                        Log-Message "${employeeName}:   Mailbox is already a Shared Mailbox"
                    }
        
                    # Log all current permissions for debugging
                    $allPermissions = Get-MailboxPermission -Identity $employeeEmail

                    # Check if the manager already has full access to Shared Mailbox
                    $permissions = $allPermissions | Where-Object { $_.User -like $managerEmail -and $_.AccessRights -contains "FullAccess" }

                    if ($permissions.Count -eq 0) {
                        # Add manager as full access delegate
                        Add-MailboxPermission -Identity $employeeEmail -User $managerUserID -AccessRights FullAccess -InheritanceType All
                        Log-Message "${employeeName}:   $managerName has been granted Full-Access to the Shared Mailbox."
                        $userSummary.SharedMailboxSet = "X"
                    } else {
                        Log-Message "${employeeName}:   $managerName already has full access to the Shared Mailbox"
                        $userSummary.SharedMailboxSet = "P"
                    }
                } catch {
                    Log-Message "${employeeName}:   Shared mailbox setup Failed."
                    Log-Message "An error occurred: $_"
                    $userSummary.SharedMailboxSet = "E"
                }
            }
        
        } elseif ($currentDateTime -ge $endMailForwarding) {
            if ($emailPreference -eq "Forwarding") {
            try {            
            # Check if forwarding needs to be removed
                        if ($currentForwardingAddress -ne $null -or $currentForwardingEnabled -ne $false) {
                            # Remove email forwarding
                            Set-Mailbox -Identity $employeeEmail -ForwardingSMTPAddress $null -DeliverToMailboxAndForward $false
                            Log-Message "${employeeName}:   Forwarding to Manager: $managerName - Ended"
                            $userSummary.MailForwardingEnded = "X"
                        } else {
                            Log-Message "${employeeName}:   Forwarding already Ended"
                            $userSummary.MailForwardingEnded = "P"
                        }
                    
                } catch {
                    Log-Message "${employeeName}:   Forwarding setup Failed"
                    Log-Message "Forwarding error occurred: $_"
                    $userSummary.MailForwardingEnded = "E"
                }
            }
            <#} elseif ($currentDateTime -ge $endMailForwarding) {    <----- Commented this out. As I don't know if we want to remove delegate access to shared mailboxes.
                    # Check if the mailbox is still shared
                    $mailbox = Get-Mailbox -Identity $employeeEmail
                    if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
                        # Optionally, you can convert it back to a regular mailbox or take other actions
                        # Set-Mailbox -Identity $employeeEmail -Type Regular
                        Log-Message "Mailbox for $employeeEmail is still shared after $endMailForwarding."
                    }
            }#>
        }


        ############ End Restructure - based on time
        
            <### OFFICE365 - Mail Forwarding - CONFIRMED WORKING
            if ($emailPreference -eq "Forwarding") {
                try {
                    # Get the current forwarding status
                    $mailbox = Get-Mailbox -Identity $employeeEmail
                    $currentForwardingAddress = $mailbox.ForwardingSMTPAddress
                    $currentForwardingEnabled = $mailbox.DeliverToMailboxAndForward
                    
                    # Log current forwarding status for debugging
                    #Log-Message "Current Forwarding Address: $currentForwardingAddress"
                    #Log-Message "Current Forwarding Enabled: $currentForwardingEnabled"
            
                    if ($currentDateTime -lt $endMailForwarding) {
                        # Check if forwarding is already set up correctly
                        if ($currentForwardingAddress -ne "smtp:$managerEmail" -or $currentForwardingEnabled -ne $true) {
                            # Set up email forwarding
                            Set-Mailbox -Identity $employeeEmail -ForwardingSMTPAddress $managerEmail -DeliverToMailboxAndForward $true
                            Log-Message "${employeeName}:   Forwarding to Manager: $managerName - Started"
                            $userSummary.MailForwardingStarted = $true
                        } else {
                            Log-Message "${employeeName}:   Forwarding already set up to $managerName."
                        }
                    } elseif ($currentDateTime -ge $endMailForwarding) {
                        # Check if forwarding needs to be removed
                        if ($currentForwardingAddress -ne $null -or $currentForwardingEnabled -ne $false) {
                            # Remove email forwarding
                            Set-Mailbox -Identity $employeeEmail -ForwardingSMTPAddress $null -DeliverToMailboxAndForward $false
                            Log-Message "${employeeName}:   Forwarding to Manager: $managerName - Ended"
                            $userSummary.MailForwardingEnded = $true
                        } else {
                            Log-Message "${employeeName}:   Forwarding already Ended"
                            $userSummary.MailForwardingEnded = $true
                        }
                    }
                } catch {
                    Log-Message "${employeeName}:   Forwarding setup Failed"
                    Log-Message "An error occurred: $_"
                }

        ### OFFICE365 - Shared Mailbox - CONFIRMED WORKING
            ## Convert separated employee mailbox to shared mailbox & add manager as full access delegate ##
            } elseif ($emailPreference -eq "Shared Mailbox") {
                try {
                    if ($currentDateTime -lt $endMailForwarding) {
                        # Check if employee's mailbox is already converted to Shared Mailbox
                        $mailbox = Get-Mailbox -Identity $employeeEmail
                        Log-Message "Current Mailbox Type: $($mailbox.RecipientTypeDetails)"
            
                        if ($mailbox.RecipientTypeDetails -ne "SharedMailbox") {
                            # If not, then Converting to shared mailbox
                            Set-Mailbox -Identity $employeeEmail -Type Shared
                            Log-Message "${employeeName}:   Converted Mailbox to a Shared Mailbox."
                        } else {
                            Log-Message "${employeeName}:   Mailbox is already a Shared Mailbox"
                        }
            
                        # Log all current permissions for debugging
                        $allPermissions = Get-MailboxPermission -Identity $employeeEmail

                        # Check if the manager already has full access to Shared Mailbox
                        $permissions = $allPermissions | Where-Object { $_.User -like $managerEmail -and $_.AccessRights -contains "FullAccess" }

                        if ($permissions.Count -eq 0) {
                            # Add manager as full access delegate
                            Add-MailboxPermission -Identity $employeeEmail -User $managerUserID -AccessRights FullAccess -InheritanceType All
                            Log-Message "${employeeName}:   $managerName has been granted Full-Access to the Shared Mailbox."
                            $userSummary.SharedMailboxSet = $true
                        } else {
                            Log-Message "${employeeName}:   $managerName already has full access to the Shared Mailbox"
                        }
                    } <#elseif ($currentDateTime -ge $endMailForwarding) {    <----- Commented this out. As I don't know if we want to remove delegate access to shared mailboxes.
                        # Check if the mailbox is still shared
                        #$mailbox = Get-Mailbox -Identity $employeeEmail
                        #if ($mailbox.RecipientTypeDetails -eq "SharedMailbox") {
                            # Optionally, you can convert it back to a regular mailbox or take other actions
                            # Set-Mailbox -Identity $employeeEmail -Type Regular
                            #Log-Message "Mailbox for $employeeEmail is still shared after $endMailForwarding."
                        #}
                    }#
                } catch {
                    Log-Message "${employeeName}:   Shared mailbox setup Failed."
                    Log-Message "An error occurred: $_"
                }
            }#>
        
                
                ##### If All Necessary Actions Completed Add "Yes" to Processed_1 Column ----Needs Work!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                
                # Get the Excel file content
                $headers = @{
                    Authorization = "Bearer $global:AccessToken"
                }
                $response = Invoke-RestMethod -Uri $global:OneDriveUrl -Headers $headers -Method Get

                # Convert the response to a usable format
                $global:data = @()
                $columns = $response.values[0]
                for ($i = 1; $i -lt $response.values.Count; $i++) {
                    $row = $response.values[$i]
                    $obj = @{}
                    for ($j = 0; $j -lt $columns.Count; $j++) {
                        $obj[$columns[$j]] = $row[$j]
                    }
                    $global:data += $obj
                }

                # Check if all actions are completed
                $allActionsCompleted = $userSummary.ADDisabled -and $userSummary.HiddenFromAddressBook -and $userSummary.OfficeSignInBlocked -and $userSummary.OfficePasswordUpdated -and $userSummary.EmailAutoReplySet -and $userSummary.SharedMailboxSet -and $userSummary.ZoomAccountDeactivated

                if ($allActionsCompleted) {
                    # Find the row for the specific user
                    $userRow = $global:data | Where-Object { $_.EmployeeEmail -eq $employeeEmail }

                    if ($userRow) {
                        # Update the Processed_1 column to "Yes"
                        $userRow.Processed_1 = "Yes"

                        # Define the URL to update the Excel file content
                        $global:OneDriveUpdateUrl = "https://graph.microsoft.com/v1.0/users/$global:OneDriveUserId/drive/items/$global:OneDriveFileId/workbook/worksheets/$global:WorksheetName/range(address='Processed_1!A1')"

                        # Prepare the body for the update request
                        $updateBody = @{
                            values = @(
                                @("Yes")
                            )
                        } | ConvertTo-Json

                        # Update the Excel file content
                        Invoke-RestMethod -Uri $global:OneDriveUpdateUrl -Headers @{
                            Authorization = "Bearer $global:AccessToken"
                            "Content-Type" = "application/json"
                        } -Method Patch -Body $updateBody
                    }
                }
        # After processing each user
        $userSummaries += $userSummary
        # Log the actions
        Log-Message "${employeeName}:   Action sequence Completed"
        Log-Message "************************************"
        }
    }
}

Log-Message "Ending Employee Separation Script..."

### Email Summary Portion #### - Confirmed Working. Need to figure out though how to handle users that have already been processed.
# Define the directory where log files are stored
$logDirectory = "\\...\SeparationScriptLogs"  # Replace with the actual path to your log files

# Get the most recent log file
$latestLogFile = Get-ChildItem -Path $logDirectory -Filter *.txt | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Read the content of the log file
$logFileContent = Get-Content -Path $latestLogFile.FullName -Raw

# Define CSS styles
$cssStyles = @"
<style>
    .centered {
        text-align: center;
        vertical-align: middle;
    }
</style>
"@

# Generate HTML table with CSS classes
$htmlTable = "$cssStyles<table border='1'><tr><th>Name</th><th>Separation Date</th><th>AD Disabled</th><th>Hidden From Address Book</th><th>Office Sign-In Blocked</th><th>Office Password Updated</th><th>Office License Updated</th><th>Email Auto-Reply Set</th><th>Mail Forwarding Started</th><th>Mail Forwarding Ended</th><th>Shared Mailbox Set</th><th>Zoom Account Deactivated</th><th>Zoom Data Transferred</th></tr>"

foreach ($user in $userSummaries) {
    $htmlTable += "<tr>"
    $htmlTable += "<td class='centered'>$($user.Name)</td>"
    $htmlTable += "<td class='centered'>$($user.SeparationDate)</td>"
    $htmlTable += "<td class='centered'>$($user.ADDisabled)</td>"
    $htmlTable += "<td class='centered'>$($user.HiddenFromAddressBook)</td>"
    $htmlTable += "<td class='centered'>$($user.OfficeSignInBlocked)</td>"
    $htmlTable += "<td class='centered'>$($user.OfficePasswordUpdated)</td>"
    $htmlTable += "<td class='centered'>$($user.OfficeLicenseUpdated)</td>"
    $htmlTable += "<td class='centered'>$($user.EmailAutoReplySet)</td>"
    $htmlTable += "<td class='centered'>$($user.MailForwardingStarted)</td>"
    $htmlTable += "<td class='centered'>$($user.MailForwardingEnded)</td>"
    $htmlTable += "<td class='centered'>$($user.SharedMailboxSet)</td>"
    $htmlTable += "<td class='centered'>$($user.ZoomAccountDeactivated)</td>"
    $htmlTable += "<td class='centered'>$($user.ZoomDataTransferred)</td>"
    $htmlTable += "</tr>"
}

$htmlTable += "</table>"

# Compose the email
$emailBody = @"
<html>
<body>
<p>Aloha Team,</p>
<p>Please find below the summary of the Employee Separation Script:</p>
$htmlTable
<p>X = Action Performed<br>P = Previously Performed<br>E = Error (See Attached Logs)<br>Blank = Not Performed Yet<br></p>
<p>Thank you,<br>Your Friendly IT AI Bot</p>
</body>
</html>
"@

# Define the email parameters
$emailParams = @{
    Message = @{
        subject = "User Action Summary"
        body = @{
            contentType = "HTML"
            content = $emailBody
        }
        toRecipients = @(@{emailAddress = @{address = "[YourUPN]"}})
        attachments = @(@{
            "@odata.type" = "#microsoft.graph.fileAttachment"
            name = $latestLogFile.Name
            contentBytes = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($logFileContent))
        })
    }
    saveToSentItems = "true"
}

# Convert the email parameters to JSON
$jsonEmailParams = $emailParams | ConvertTo-Json -Depth 10

# Send the email
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/[ReportRecipientUPN]/sendMail" -Method Post -Headers @{
    Authorization = "Bearer $global:AccessToken"
    "Content-Type" = "application/json"
} -Body $jsonEmailParams




$userSummaries = @()
Disconnect-MgGraph
exit 0 #Success code for Task Scheduler
