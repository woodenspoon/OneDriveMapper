<#
.SYNOPSIS

Save the Teams SharePoint library IDs into an Azure App Function.

.DESCRIPTION

This function will connect to a tenant and find the Azure App Function matching the pattern 'OneDriveMapper*'.
Once found, the function will read the CSV file containing the library IDs and save a JSON representation into the Variable name 'OneDriveSyncUrls'.
This information is used by the Azure App Function to return all the authorized Library IDs for a particular user when called at logon time.

.PARAMETER CSVFile

The CSV file to export the library IDs to, defaults to '.\OneDriveSyncUrls.csv'

#>
Function Publish-SharePointLibraryIDs {

    [CmdletBinding()]
    Param(
        [System.IO.FileInfo]$CSVFile = '.\OneDriveSyncUrls.csv'
    )
    
    Import-Module Az.Accounts -ErrorAction Stop
    Import-Module Az.Functions -ErrorAction Stop
    
    try {
    
        $urls = Import-Csv $CSVFile
    
        Write-Host "Connecting to Azure. Please sign-in with a Global Admin user for the target tenant in the next window." -ForegroundColor Green
        $azure = Connect-AzAccount
        $loggedInAdmin = $azure.Context.Account.Id
        $subscriptionName = $azure.Context.Subscription.Name
        Write-Host "Connected to Azure subscription '$subscriptionName' as $loggedInAdmin." -ForegroundColor Cyan
    
        Write-Host "Trying to retrieve Function App 'OneDriveMapper*'." -ForegroundColor Green
        $azapp = Get-AzFunctionApp | Where-Object Name -like 'OneDriveMapper*'
        Write-Host "Found app '$($azapp.Name)'." -ForegroundColor Cyan
    
        Write-Host "Setting 'OneDriveSyncUrls' value from CSV content." -ForegroundColor Green
        $odsujs = $urls | ConvertTo-Json -Compress
        $azapp | Update-AzFunctionAppSetting -AppSetting @{'OneDriveSyncUrls' = $odsujs} | Out-Null
        Write-Host "Completed." -ForegroundColor Cyan
    
    } catch {
    
        throw $_
    
    }
}

<#
$azsettings = $azapp | Get-AzFunctionAppSetting
$urls = $azsettings['OneDriveSyncUrls'] | ConvertFrom-Json
#>
