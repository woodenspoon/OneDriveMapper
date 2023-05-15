<#
.SYNOPSIS

Mounts the Teams SharePoint libraries allowed for the current OneDrive user for syncing them with OneDrive

.DESCRIPTION

This function will interrogate the GetOneDriveTeamMappings Azure Function App on the tenant and mounts the Teams SharePoint libraries that that user is allowed to sync.

#>
Function Mount-UserOneDriveTeamMappings {
    [CmdletBinding()]
    Param(
        [ValidatePattern('^[a-zA-Z0-9]+$')][String]$OneDriveMapperFunctionApp,
        [ValidatePattern('^[A-Za-z0-9+/=]|=[^=]|={3,}$')][String]$OneDriveMapperAuthKey,
        [String]$UserEmail,
        [switch]$DryRun
    )

    # Derive the parameters if they were not explicitely passed

    # OneDrive for Business signed-in user
    if (-not $userEmail) {
        $userEmail = Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserEmail
    }
    if (-not $userEmail) {
        throw "Unable to determine OneDrive for Business user, make sure the user is signed in to OneDrive, aborting."
    }

    # Function app name, stored in the registry using Set-UserOneDriveTeamMappingSettings cmdlet
    if (-not $OneDriveMapperFunctionApp) {
        $OneDriveMapperFunctionApp = Get-ItemProperty "HKLM:\SOFTWARE\WST\OneDriveMapper" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OneDriveMapperFunctionApp
    }
    if (-not $OneDriveMapperFunctionApp) {
        throw "Unable to determine OneDriveMapperFunctionApp name, run Set-UserOneDriveTeamMappingSettings to setup, aborting."
    }

    # Function auth key, stored in the registry using Set-UserOneDriveTeamMappingSettings cmdlet
    if (-not $OneDriveMapperAuthKey) {
        $OneDriveMapperAuthKey = Get-ItemProperty "HKLM:\SOFTWARE\WST\OneDriveMapper" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OneDriveMapperAuthKey
    }
    if (-not $OneDriveMapperAuthKey) {
        throw "Unable to determine OneDriveMapperAuthKey value, run Set-UserOneDriveTeamMappingSettings to setup, aborting."
    }

    # Announce which user is being considered
    Write-Host "Attempting to mount Teams SharePoint libraries in OneDrive for user '$userEmail'" -ForegroundColor DarkBlue -BackgroundColor Cyan

    # Query the OneDriveMapper function app to retrieve the list of Teams information blocks
    [hashtable]$postCallResult = @{}
    $postCallResult.error = $false
    $url = "https://$OneDriveMapperFunctionApp.azurewebsites.net/api/GetOneDriveTeamMappings"
    $headers = @{'x-functions-key' = $OneDriveMapperAuthKey}
    $body = ConvertTo-Json @{userEmail = $userEmail}

    try {

        $results = Invoke-WebRequest -Uri $url -Method Post -Body $body -Headers $headers -ContentType "application/json" -UseBasicParsing

        # Process the response
        if ($results.StatusCode -ne 200) {
            $postCallResult.exception = @{'response' = @{'StatusCode' = $results.StatusCode; 'StatusDescription' = $results.Content}}
            $postCallResult.error = $true
        } else {
            $results = $results.Content | ConvertFrom-Json
        }

    } catch {

        $postCallResult.exception = $_.Exception
        $postCallResult.error = $true

    }

    # Something went wrong
    if ($postCallResult.error) {

        if ($postCallResult.exception.response.StatusCode) {
            Write-Error ("Error connecting: Response was {0} - {1}" -f $postCallResult.exception.response.StatusCode, $postCallResult.exception.response.StatusDescription)
        } else {
            Write-Error ("Error connecting: Response was {0:x} - {1}" -f $postCallResult.exception.HResult, $postCallResult.exception.Message)
        }

    # All good, process each entry and determine if this Teams entry should and can be synced safely
    } else {

        $results | ForEach-Object {

            # URI to sync this folder
            $channelUrl = ("userEmail={0}&siteId={{{1}}}&webId={{{2}}}&webTitle={3}&webUrl={4}&listId={{{5}}}&folderId={{{6}}}&folderName={7}&folderUrl={8}&version=1&scope=OPENFOLDER" -f `
                $userEmail, $_.SiteId, $_.WebId, $_.WebTitle, $_.WebUrl, $_.ListId, $_.FolderId, $_.ChannelTitle, $_.FolderUrl)
            $channelLaunch = "odopen://sync/?" + ($channelUrl -replace '{','%7B' -replace '}','%7D' -replace ' ', '%20' -replace ':','%3A' -replace '/','%2F')

            # Expected local path
            $sanitizedWebTitle = ($_.WebTitle -replace '[<>:"/\|?*]', ' ')
            $ODpath = ("{0}\{1}\{2} - {3}" -f $ENV:USERPROFILE, $_.TenantTitle, $sanitizedWebTitle, $_.ChannelTitle)

            # Expected registry mapping (which may have been done when the tenant name was different)
            $ODpattern = ("{0}\{1}\{2} - {3}" -f $ENV:USERPROFILE, ".*", $sanitizedWebTitle, $_.ChannelTitle) -replace '\\', '\\'
            $ODreg = Get-Item ("HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants\{0}" -f $_.TenantTitle) -ErrorAction SilentlyContinue
            if ($ODreg) {
                $ODfolder = $ODReg.Property | Where-Object {$_ -match $ODpattern}
            } else {
                $ODfolder = $null
            }

            # Decide on course of action
            if ($ODFolder) {
                if ($ODFolder -ne $ODPath) {
                    # Already synced to another folder (no action)
                    Write-Host "$($_.WebTitle) synced to DIFFERENT folder: $ODFolder" -ForegroundColor Yellow -BackgroundColor Black
                } else {
                    # Already synced (no action)
                    Write-Host "$($_.WebTitle) synced to folder: $ODFolder" -ForegroundColor DarkBlue -BackgroundColor Cyan
                }
            } else {
                if (Test-Path $ODPath) {
                    # A folder exists, cannot sync without user intervention
                    Write-Host "Re-syncing $($_.WebTitle) that was PREVIOUSLY synced to folder: $ODPath, pausing for 20 seconds" -ForegroundColor Yellow -BackgroundColor Black
                    if (-not $DryRun) {
                        Start-Process $channelLaunch
                        Start-Sleep -Seconds 20
                    }
                } else {
                    # Sync folder
                    Write-Host "Syncing $($_.WebTitle) to folder: $ODPath, pausing for 20 seconds" -ForegroundColor DarkGreen -BackgroundColor Cyan
                    if (-not $DryRun) {
                        Start-Process $channelLaunch
                        Start-Sleep -Seconds 20
                    }
                }
            }

        }

    }

}
