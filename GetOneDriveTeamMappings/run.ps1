<#

Function App URL: https://onedrivemapper4827.azurewebsites.net/api/GetOneDriveTeamMappings?code=BhQq7isH5jj0WCcxtbQSrdkLjS29jWOdQvPnabj95zLaAzFujWgOKw==

#>

using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Request can be either a GET or a POST, action parameter can come as either
$userEmail = $Request.Query.userEmail
if (-not $userEmail) {
    $userEmail = $Request.Body.userEmail
}

$body = ConvertTo-Json(@{ReturnCode = 0; Message = "This HTTP triggered function executed successfully. Pass userEmail in the query string to request the OneDrive Teams mapping information for that user."})

if ($userEmail) {
    $body = ConvertTo-Json(@{ReturnCode = 0; Message = "Returning OneDrive Teams mapping information for $userEmail."})
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
