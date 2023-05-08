using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$userEmail = $Request.Query.userEmail
if (-not $userEmail) {
    $userEmail = $Request.Body.userEmail
}

# Process the request
try {

    # Retrieve the valid OneDrive URLs
    $OneDriveSyncUrls = $ENV:OneDriveSyncUrls
    if (-not $OneDriveSyncUrls) {
        throw "Azure Function App environment variable 'OneDriveSyncUrls' missing."
    }
    $OneDriveSyncUrls = ($ENV:OneDriveSyncUrls).Value | ConvertFrom-Json

    # Retrieve the user email
    if (-not $userEmail) {
        throw "Missing userEmail parameter"
    }

    # Determine the groups this user belongs to
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users
    Connect-MgGraph -AccessToken ((Get-AzAccessToken -ResourceTypeName MSGraph).token)
    $userGroups = Get-MgUserMemberOf -UserId $userEmail

    # If the user is found, get the ID of the groups it belongs to
    if (-not $userGroups) {
        throw "Unknown user '$userEmail'"
    }
    $userGroupsIDs = $userGroups | Where-Object {$null -eq $_.DeletedDateTime} | Select-Object -ExpandProperty Id

    $body = $userGroupsIDs | ConvertTo-Json
    $retCode = [HttpStatusCode]::OK

} catch {

    # Report errors back to the caller
    $body = $_.Exception.Message
    $retCode = [HttpStatusCode]::NonAuthoritativeInformation

}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = $retCode
    Body = $body
})
