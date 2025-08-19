function Update-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )

    # Get the installed module
    $installedModule = Get-InstalledModule -Name $ModuleName -ErrorAction SilentlyContinue

    if ($null -ne $installedModule) {
        # Find the latest version of the module
        $latestModule = Find-Module -Name $ModuleName

        if ($latestModule.Version -gt $installedModule.Version) {
            Write-Host "Updating module $ModuleName from version $($installedModule.Version) to $($latestModule.Version)"
            Update-Module -Name $ModuleName
        } else {
            Write-Host "Module $ModuleName is up to date (version $($installedModule.Version))"
        }
    } else {
        Write-Host "Module $ModuleName is not installed. Installing the latest version."
        Install-Module -Name $ModuleName
    }
}

# List of modules to check and update
$modulesToCheck = @("ExchangeOnlineManagement", "CredentialManager", "Microsoft.Graph", "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Files", "ImportExcel")

foreach ($module in $modulesToCheck) {
    Update-ModuleIfNeeded -ModuleName $module
}
