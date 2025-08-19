# This script automates the creation of room calendars in Exchange Online using room names from an Excel file.

# Install and import Exchange Online module
# Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName [YourUPN]

# Path to the Excel file
$excelFilePath = "\\...\ConferenceRoomImport.xlsx"

# Load the Excel file
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelFilePath)
$sheet = $workbook.Sheets.Item(1)

# Get the range of used cells
$usedRange = $sheet.UsedRange
$roomNames = @()

# Read room names from the Excel file
for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
    $roomName = $sheet.Cells.Item($row, 1).Text
    $roomNames += $roomName
}

# Close the Excel file
$workbook.Close($false)
$excel.Quit()

# Create room calendars in Exchange Online with custom names and email addresses
foreach ($roomName in $roomNames) {
    $customRoomName = "LC $roomName"
    $emailAddress = "LC$($roomName -replace ' ', '')@[EmailDomain]"
    New-Mailbox -Name $customRoomName -Room -PrimarySmtpAddress $emailAddress
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
