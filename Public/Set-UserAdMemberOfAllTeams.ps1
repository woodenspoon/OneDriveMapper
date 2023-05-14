<#
.SYNOPSIS

Adds a Global Admin to all Teams.

.DESCRIPTION

This function will login to a tenant and assign the logged in user to all the Teams as either a Member or an Owner.
This is required to allow that user to then fetch the SharePoint details for those Teams in order to generate the OneDrive map for syncing them.

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
        
        # Make sure we are at least a member on all the Teams
        Write-Host "Connecting to Microsoft Team. Please re-use the previous authentication credentials if offered." -ForegroundColor Green
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host "Connected." -ForegroundColor Cyan
        
        Write-Host "Adding $loggedInAdmin with Owner role on all tenant's Teams" -ForegroundColor Green
        Get-Team | Add-TeamUser -User $loggedInAdmin -Role $Role
        Write-Host "Completed." -ForegroundColor Cyan
        
    } catch {

        throw $_

    }
}
