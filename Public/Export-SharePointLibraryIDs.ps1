<#
.SYNOPSIS

Export to a CSV all Teams SharePoint library IDs necessary for syncing them with OneDrive

.DESCRIPTION

This function will connect to a tenant and find all the Teams SharePoint libraries it can access and generate the data necessary to sync those Teams via OneDrive using the odopen: url.
During execution, the function will ask the operator to sign-in to Microsoft 365. A Global Admin is recommended for this sign-in.

.PARAMETER DocLibraryName

The SharePoint library name of the library contains the Teams files, defaults to 'Documents'

.PARAMETER DocLibraryTitle

The SharePoint library title of the library contains the Teams files, defaults to 'Shared Documents'

.PARAMETER ChannelTitle

The default Team channel for all Teams, defaults to 'General'

.PARAMETER CSVFile

The CSV file to export the library IDs to, defaults to '.\OneDriveSyncUrls.csv'

#>
Function Export-SharePointLibraryIDs {
    [CmdletBinding()]
    Param(
        [String]$DocLibraryName = 'Documents',
        [String]$DocLibraryTitle = 'Shared Documents',
        [String]$ChannelTitle = 'General',
        [System.IO.FileInfo]$CSVFile = '.\OneDriveSyncUrls.csv'
    )

    # Load necessary modules
    Import-Module AzureAD -ErrorAction Stop
    Import-Module PnP.PowerShell -ErrorAction Stop

    try {

        # Get tenant information
        Write-Host "Connecting to Azure AD. Please sign-in with a Global Admin user for the target tenant in the next window." -ForegroundColor Green
        $aad = Connect-AzureAD -ErrorAction Stop
        $loggedInAdmin = $aad.Account.Id
        $tenant = (Get-AzureADDomain | Where-Object Name -Match '^[a-z-]+\.onmicrosoft\.com' | Select-Object -ExpandProperty Name) -replace '\.onmicrosoft\.com'
        $tenantTitle = Get-AzureADTenantDetail | Select-Object -ExpandProperty DisplayName
        Write-Host "Connected to $tenant.onmicrosoft.com ($tenantTitle) as $loggedInAdmin." -ForegroundColor Cyan

        # Get all the Teams' SharePoint site information
        Write-Host "Connecting to SharePoint Online. Please re-use the previous authentication credentials if offered." -ForegroundColor Green
        Connect-PnPOnline -Url https://$tenant.sharepoint.com -Interactive -ErrorAction Stop | Out-Null
        Write-Host "Connected." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Getting all Teams' SharePoint information." -ForegroundColor Green
        $teams = Get-PnPTeamsTeam | Sort-Object -Property DisplayName
        Write-Host "Done." -ForegroundColor Cyan
        Write-Host ""

        Write-Host "Start generating Library IDs for each Team." -ForegroundColor Green
        $urls = @()
        $cnt = 0
        foreach ($team in $teams) {
            $webTitle = $team.DisplayName
            $groupId = $team.GroupId
            Write-Progress -Activity "Retrieving SharePoint information" -Status "Reading Team $webTitle" -PercentComplete ($cnt++ / $teams.count * 100)
            Write-Host "..Processing Team $webTitle" -ForegroundColor Cyan

            # Verify we can access the SharePoint information for this Team (any Owner or Member can)
            if (Get-AzureADGroupMember -ObjectId $team.groupid | Where-Object UserPrincipalName -eq $loggedInAdmin) {

                $webUrl = Get-PnPMicrosoft365Group -IncludeSiteUrl -Identity $groupId | Select-Object -ExpandProperty SiteUrl
                Connect-PnPOnline $webUrl -Interactive -ErrorAction Stop | Out-Null
                $siteId = Get-PnPSite -Includes Id -ErrorAction Stop | Select-Object -ExpandProperty Id
                $webId = Get-PnPWeb -Includes Id -ErrorAction Stop | Select-Object -ExpandProperty Id
                $listId = (Get-PnPList $DocLibraryName -Includes Id -ErrorAction Stop | Select-Object -ExpandProperty Id).ToString().ToUpper()
                $folder = Get-PnPFolderItem -FolderSiteRelativeUrl ('/{0}' -f $DocLibraryTitle) -ErrorAction Stop | Where-Object Name -eq $ChannelTitle | Select-Object -First 1
                $folderId = $folder.UniqueId
                $folderUrl = "{0}/{1}/{2}" -f $webUrl, $DocLibraryTitle, $ChannelTitle
        
                # Return the relevant information to allow the caller to build the OD sync URI
                $urls += [PSCustomObject]@{
                    TenantTitle = $tenantTitle
                    WebTitle = $webTitle
                    ChannelTitle = $ChannelTitle
                    GroupId = $groupId
                    SiteId = $siteId
                    WebId = $webId
                    ListId = $listId
                    FolderId = $folderId
                    WebUrl = $webUrl
                    FolderUrl = $folderUrl
                }

            } else {

                Write-Warning "You ($loggedInAdmin) are not a member of '$webTitle', skipping"

            }
        }

        Write-Host "Done." -ForegroundColor Cyan
        $urls | Export-Csv -NoTypeInformation -Path $CSVFile
        Write-Host ""
        Write-Host "Exported LibraryIDs to CSV file '$CSVFile'." -ForegroundColor Cyan
        Write-Host "Operation completed successfully." -ForegroundColor Cyan

    } catch {

        throw $_

    }

}
