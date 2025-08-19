# This script performs reverse IP lookups using ipinfo.io and updates an Excel file with location data for each IP address.

function Module_Setup {
    try {
        Import-Module ImportExcel
    } catch {
        Write-Host "An error occurred: $_"
    }
}

Module_Setup

# Function to validate IP addresses
function Validate-IP {
    param ([string]$ipAddress)

    $ipv4Pattern = '^(\d{1,3}\.){3}\d{1,3}$'
    $ipv6Pattern = '^([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}$'

    return ($ipAddress -match $ipv4Pattern) -or ($ipAddress -match $ipv6Pattern)
}

# Function to perform reverse lookup on IP addresses using ipinfo.io
function Get-LocationFromIP {
    param (
        [string]$ipAddress,
        [string]$apiKey
    )

    $url = "https://ipinfo.io/${ipAddress}?token=${apiKey}"
    Write-Host "Requesting URL: $url"
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        return $response
    } catch {
        Write-Host "Failed to lookup IP address: $ipAddress. Error: $_"
        return $null
    }
}

# Function to update Excel with location information
function Update-ExcelWithLocation {
    param (
        [string]$inputExcelPath,
        [string]$outputExcelPath,
        [string]$apiKey
    )

    $data = Import-Excel -Path $inputExcelPath

    $columnsToAdd = @("IPCity", "IPRegion", "IPCountry", "IPASN", "IPISP", "IPLatitude", "IPLongitude")
    foreach ($column in $columnsToAdd) {
        $data | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name $column -Value "" }
    }

    $ipCache = @{}
    $invalidIPs = @()

    foreach ($row in $data) {
        $ipAddress = $row.IPAddress
        Write-Host "Processing IP address: $ipAddress"

        if ([string]::IsNullOrEmpty($ipAddress)) {
            Write-Host "IP address is null or empty for row: $($row | Out-String)"
            continue
        }

        if (-not (Validate-IP -ipAddress $ipAddress)) {
            $invalidIPs += $ipAddress
            continue
        }

        if ($ipCache.ContainsKey($ipAddress)) {
            $location = $ipCache[$ipAddress]
        } else {
            $retryCount = 0
            $maxRetries = 3
            $location = $null

            while ($retryCount -lt $maxRetries -and $location -eq $null) {
                $location = Get-LocationFromIP -ipAddress $ipAddress -apiKey $apiKey
                if ($location -eq $null) {
                    $retryCount++
                    Write-Host "Retrying lookup for IP address: $ipAddress ($retryCount/$maxRetries)"
                    Start-Sleep -Seconds 2
                }
            }

            if ($location -ne $null) {
                $ipCache[$ipAddress] = $location
            } else {
                Write-Host "Failed to lookup IP address after retries: $ipAddress"
                continue
            }
        }

        $row.IPCity      = $location.city
        $row.IPRegion    = $location.region
        $row.IPCountry   = $location.country
        $row.IPASN       = $location.org
        $row.IPISP       = $location.org
        $coordinates     = $location.loc -split ","
        $row.IPLatitude  = $coordinates[0]
        $row.IPLongitude = $coordinates[1]
    }

    $data | Export-Excel -Path $outputExcelPath -AutoSize
    Write-Host "The new Excel file has been created with location information."

    if ($invalidIPs.Count -gt 0) {
        Write-Host "The following IP addresses were invalid and skipped:"
        $invalidIPs | ForEach-Object { Write-Host $_ }
    }
}

# Example usage
$inputExcelPath = "\\...\PS\IPReverseLookupFiles\nonHI_30days_sign_in_logs.xlsx"
$outputExcelPath = "\\...\NEW_nonHI_30days_sign_in_logs.xlsx"
$apiKey = "[YourIPInfoAPIKey]"

Update-ExcelWithLocation -inputExcelPath $inputExcelPath -outputExcelPath $outputExcelPath -apiKey $apiKey
