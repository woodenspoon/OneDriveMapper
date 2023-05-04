<#
.SYNOPSIS

Return the Teams SharePoint library IDs allowed for the current OneDrive user for syncing them with OneDrive

.DESCRIPTION

This function will interrogate the GetOneDriveTeamMappings Azure Function App on the tenant and return the Teams SharePoint library IDs that that user is allowed to sync.

#>
Function Get-UserOneDriveTeamMappings {
    [CmdletBinding()]
    Param(
        [ValidatePattern('^[a-zA-Z0-9]+$')][Parameter(Mandatory=$true)][String]$OneDriveMapperFunctionApp,
        [ValidatePattern('^[A-Za-z0-9+/=]|=[^=]|={3,}$')][Parameter(Mandatory=$true)][String]$OneDriveMapperAuthKey,
        [Parameter(Mandatory=$true)][String]$userEmail
    )


    [hashtable]$postCallResult = @{}
    $postCallResult.error = $false
    $url = "https://$OneDriveMapperFunctionApp.azurewebsites.net/api/GetOneDriveTeamMappings"
    $headers = @{'x-functions-key' = $OneDriveMapperAuthKey}
    $body = ConvertTo-Json @{userEmail = $userEmail}

    try {

        $results = Invoke-RestMethod -Uri $url -Method Post -Body $body -Headers $headers -ContentType "application/json"

    } catch {

        $postCallResult.exception = $_.Exception
        $postCallResult.error = $true

    }

    if ($postCallResult.error) {

        #$postCallResult | ConvertTo-Json -Depth 2
        if ($postCallResult.exception.response.StatusCode) {
            Write-Error ("Error connecting: Response was {0} - {1}" -f $postCallResult.exception.response.StatusCode, $postCallResult.exception.response.StatusDescription)
        } else {
            Write-Error ("Error connecting: Response was {0:x} - {1}" -f $postCallResult.exception.HResult, $postCallResult.exception.Message)
        }

    } else {

        $results

    }

}
