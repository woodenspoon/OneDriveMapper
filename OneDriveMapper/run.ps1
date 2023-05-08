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

$body = "This HTTP triggered function executed successfully. Pass a userEmail in the query string or in the request body for a personalized response."

if ($userEmail) {
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users
    Connect-MgGraph -AccessToken ((Get-AzAccessToken -ResourceTypeName MSGraph).token)
    $body = Get-MgUserMemberOf -UserId $userEmail | ConvertTo-Json
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
