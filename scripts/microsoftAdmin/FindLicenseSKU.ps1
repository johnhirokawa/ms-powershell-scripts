# This script connects to Microsoft Graph to retrieve Copilot SKU details.

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Organization.Read.All"

# Retrieve the list of subscribed SKUs
$SubscribedSkus = Get-MgSubscribedSku

# Find the SKU for "Copilot for Microsoft 365"
$CopilotSku = $SubscribedSkus | Where-Object { $_.SkuPartNumber -eq "[CopilotSkuPartNumber]" } | Select-Object SkuId, SkuPartNumber

# Display the SKU ID for "Copilot for Microsoft 365"
$CopilotSku
