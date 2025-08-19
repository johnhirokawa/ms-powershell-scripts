# Undo PAC file enforcement for Edge, Chrome, and Firefox

# Remove Edge PAC settings
Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Edge" -Name "ProxyMode" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Edge" -Name "ProxyPacUrl" -ErrorAction SilentlyContinue

# Remove Chrome PAC settings
Remove-ItemProperty -Path "HKLM:\Software\Policies\Google\Chrome" -Name "ProxyMode" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKLM:\Software\Policies\Google\Chrome" -Name "ProxyPacUrl" -ErrorAction SilentlyContinue

# Remove Firefox policies.json
$firefoxPolicyPath = "C:\Program Files\Mozilla Firefox\distribution\policies.json"
if (Test-Path $firefoxPolicyPath) {
    Remove-Item $firefoxPolicyPath -Force
    Write-Output "Deleted Firefox policy file: $firefoxPolicyPath"
} else {
    Write-Output "Firefox policy file not found: $firefoxPolicyPath"
}

# Optional: Refresh Group Policy
gpupdate /force
