##Cert Install
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process
# Expected SHA256 hash of the certificate
$expectedHash = "abcde1234a;ldjfalskjflsadkjfdl;sajfl;asdjf"  # Replace with actual hash

# Download the certificate
Invoke-WebRequest -Uri "[insert download link for cert file]" -OutFile "$env:TEMP\cloud_gateway.cer"

# Compute hash
$actualHash = Get-FileHash "$env:TEMP\cloud_gateway.cer" -Algorithm SHA256 | Select-Object -ExpandProperty Hash

# Compare and install
if ($actualHash -eq $expectedHash) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import("$env:TEMP\cloud_gateway.cer")
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
    $store.Open("ReadWrite")
    $store.Add($cert)
    $store.Close()
    Write-Host "Certificate installed successfully." -ForegroundColor Green
} else {
    Write-Host "Certificate hash mismatch. Aborting installation." -ForegroundColor Red
}
