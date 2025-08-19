# This script queries Microsoft Graph for sign-in logs, filters for non-Hawaii activity, and exports results to Excel for CSEA discovery.

Write-Output "Starting CSEA Discovery Script..."

## Idea: Consider adding functionality to disable user in Mosyle (MDM platform)
## Description: PowerShell Script - Email Forwarding & Auto Reply
## Client ID: [ClientID]
## Tenant ID: [TenantID]

# NOTE: Microsoft.Graph module installation takes time and should only be done once per machine.

<# 
Install-Module ActiveDirectory -Force
Install-Module ExchangeOnlineManagement -Force
Install-Module -Name PnP.PowerShell -Scope CurrentUser
Install-Module CredentialManager -Force
Install-Module -Name Microsoft.Graph -Force
Install-Module -Name ImportExcel -Scope CurrentUser
#>

function Module_Setup {
    try { Import-Module ImportExcel } catch { Log-Message "An error occurred: $_" }
    try { Import-Module ActiveDirectory } catch { Log-Message "An error occurred: $_" }
    try { Import-Module ExchangeOnlineManagement } catch { Log-Message "An error occurred: $_" }
    try { Import-Module CredentialManager } catch { Log-Message "An error occurred: $_" }
    try { Import-Module Microsoft.Graph.Authentication } catch { Log-Message "An error occurred: $_" }
    try { Import-Module Microsoft.Graph.Users } catch { Log-Message "An error occurred: $_" }
    try { Import-Module Microsoft.Graph.Files } catch { Log-Message "An error occurred: $_" }
}

Module_Setup

function Mg-GraphConnect {
    $global:ExchangeAPICred = Get-StoredCredential -Target "ExchangeAPI"
    if ($null -eq $global:ExchangeAPICred) {
        Write-Host "Failed to retrieve credentials. Please check the target name."
        return
    }

    $global:TenantId = "[TenantID]"
    $global:ClientId = $global:ExchangeAPICred.UserName
    $plainClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:ExchangeAPICred.Password)
    )
    $global:ClientSecret = ConvertTo-SecureString -String $plainClientSecret -AsPlainText -Force

    Write-Host "Client ID: $global:ClientId"

    if ([string]::IsNullOrEmpty($global:ClientId) -or [string]::IsNullOrEmpty($plainClientSecret)) {
        Write-Host "Client ID or Client Secret is null or empty. Please check your credentials."
        return
    }

    $token = Get-MsalToken -TenantId $global:TenantId -ClientId $global:ClientId -ClientSecret $global:ClientSecret -Scopes "https://graph.microsoft.com/.default"
    if ($null -eq $token) {
        Write-Host "Failed to retrieve access token."
        return
    }

    $global:AccessToken = $token.AccessToken
    Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $global:AccessToken -AsPlainText -Force)
    Write-Host "Connection Completed"
}

function ConvertTo-HawaiiTime {
    param ([datetime]$utcDateTime)
    return $utcDateTime.ToUniversalTime().AddHours(-10)
}

function Get-SignInLogs {
    param (
        [datetime]$startDate,
        [datetime]$endDate
    )

    $query = "https://graph.microsoft.com/v1.0/auditLogs/signIns?\$filter=createdDateTime ge $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and createdDateTime le $($endDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    $allSignInLogs = @()

    do {
        $retryCount = 0
        $maxRetries = 3
        $success = $false

        while (-not $success -and $retryCount -lt $maxRetries) {
            try {
                $response = Invoke-RestMethod -Uri $query -Headers @{Authorization = "Bearer $($global:AccessToken)"} -Method Get
                $success = $true
            } catch {
                $retryCount++
                Start-Sleep -Seconds ($_.Exception.Response.Headers["Retry-After"] ?? 5)
                Write-Host "Retrying... ($retryCount/$maxRetries)"
            }
        }

        if (-not $success) {
            Write-Host "Failed to retrieve sign-in logs after $maxRetries attempts."
            return
        }

        $allSignInLogs += $response.value
        $query = $response.'@odata.nextLink'
    } while ($query)

    $nonHawaiiSignIns = $allSignInLogs | Where-Object { $_.location.state -ne "Hawaii" }

    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("User")
    $dataTable.Columns.Add("Location")
    $dataTable.Columns.Add("Status")
    $dataTable.Columns.Add("DateTime (HST)")
    $dataTable.Columns.Add("IPAddress")
    $dataTable.Columns.Add("ErrorCode")

    $nonHawaiiSignIns | ForEach-Object {
        $row = $dataTable.NewRow()
        $row["User"] = $_.userDisplayName
        $row["Location"] = "$($_.location.city), $($_.location.state), $($_.location.country)"
        $row["Status"] = if ($_.status.errorCode -eq 0) { "Successful" } else { "Unsuccessful" }
        $row["DateTime (HST)"] = ConvertTo-HawaiiTime -utcDateTime $_.createdDateTime
        $row["IPAddress"] = $_.ipAddress
        $row["ErrorCode"] = $_.status.errorCode
        $dataTable.Rows.Add($row)
    }

    $dataTable | Export-Excel -Path "sign_in_logs.xlsx" -AutoSize
    Write-Host "The sign-in logs have been saved to sign_in_logs.xlsx"
}

# Example usage
$startDate = (Get-Date).AddMonths(-1)
$endDate = Get-Date

Mg-GraphConnect
Get-SignInLogs -startDate $startDate -endDate $endDate
Disconnect-MgGraph
exit 0
