# PowerShell Script to Enforce PAC File Across Edge, Chrome, and Firefox

# Set PAC file for Microsoft Edge
$edgeRegPath = "HKLM:\Software\Policies\Microsoft\Edge"
New-Item -Path $edgeRegPath -Force | Out-Null
Set-ItemProperty -Path $edgeRegPath -Name "ProxyMode" -Value "pac_script"
Set-ItemProperty -Path $edgeRegPath -Name "ProxyPacUrl" -Value "[Insert PacFile URL]" #<-- Add PacFile URL here

# Set PAC file for Google Chrome
$chromeRegPath = "HKLM:\Software\Policies\Google\Chrome"
New-Item -Path $chromeRegPath -Force | Out-Null
Set-ItemProperty -Path $chromeRegPath -Name "ProxyMode" -Value "pac_script"
Set-ItemProperty -Path $chromeRegPath -Name "ProxyPacUrl" -Value "[Insert PacFile URL]" #<-- Add PacFile URL here

# Deploy policies.json for Mozilla Firefox
$firefoxPolicyPath = "C:\Program Files\Mozilla Firefox\distribution"
New-Item -Path $firefoxPolicyPath -ItemType Directory -Force | Out-Null

$policyJson = @"
{
  "policies": {
    "Proxy": {
      "Mode": "autoConfig",
      "AutoConfigURL": "[Insert PacFile URL]"
    },
    "Preferences": {
      "network.proxy.autoconfig_url": {
        "Value": "[Insert PacFile URL]",
        "Status": "locked"
      }
    }
  }
}
"@

$policyJson | Set-Content -Path "$firefoxPolicyPath\policies.json" -Encoding UTF8

# Optional: Force Group Policy update
gpupdate /force
