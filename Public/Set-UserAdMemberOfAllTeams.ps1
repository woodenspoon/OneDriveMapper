<#
.SYNOPSIS

Save the Teams SharePoint library IDs into an Azure App Function.

.DESCRIPTION

This function will connect to a tenant and find the Azure App Function matching the pattern 'OneDriveMapper*'.
Once found, the function will read the CSV file containing the library IDs and save a JSON representation into the Variable name 'OneDriveSyncUrls'.
This information is used by the Azure App Function to return all the authorized Library IDs for a particular user when called at logon time.

.PARAMETER Role

The role to assign to the user, defaults to 'Member' but could also be set to 'Owner'

#>
Function Set-UserAdMemberOfAllTeams {

    [CmdletBinding()]
    Param(
        [String]$Role = 'Member'
    )

    Import-Module AzureAD -ErrorAction Stop
    Import-Module MicrosoftTeams -ErrorAction Stop

    try {

        Write-Host "Connecting to Azure AD. Please sign-in with a Global Admin user for the target tenant in the next window." -ForegroundColor Green
        $aad = Connect-AzureAD -ErrorAction Stop
        $loggedInAdmin = $aad.Account.Id
        $tenant = (Get-AzureADDomain | Where-Object Name -Match '^[a-z-]+\.onmicrosoft\.com' | Select-Object -ExpandProperty Name) -replace '\.onmicrosoft\.com'
        $tenantTitle = Get-AzureADTenantDetail | Select-Object -ExpandProperty DisplayName
        Write-Host "Connected to $tenant.onmicrosoft.com ($tenantTitle) as $loggedInAdmin." -ForegroundColor Cyan
        
        # Make sure we are an owner on all the Teams
        Write-Host "Connecting to Microsoft Team. Please re-used the previous authentication credentials if offered." -ForegroundColor Green
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host "Connected." -ForegroundColor Cyan
        
        Write-Host "Adding $loggedInAdmin with Owner role on all tenant's Teams" -ForegroundColor Green
        Get-Team | Add-TeamUser -User $loggedInAdmin -Role $Role
        Write-Host "Completed." -ForegroundColor Cyan
        
    } catch {

        throw $_

    }
}
