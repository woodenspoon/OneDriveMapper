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

try {
    # Logging in to Azure.
    Connect-AzAccount -Identity

    # Get token and connect to MgGraph
    Connect-MgGraph -AccessToken ((Get-AzAccessToken -ResourceTypeName MSGraph).token)
} catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

# Test if we can get users:
$body = Get-MgUser -All

<#
$body = @{ReturnCode = 1; Message = "This HTTP triggered function executed successfully. Pass userEmail in the query string to request the OneDrive Teams mapping information for that user."}

if ($userEmail) {
    $body = @{ReturnCode = 0; LibraryIDs = @(
        @{
            TenantTitle = 'WST'
            WebTitle     = 'Admin'
            ChannelTitle = 'General'
            GroupId      = 'fe4836d1-495d-43d0-9f59-b831e928c7d2'
            SiteId       = '4c15fc62-380f-4f05-8792-95d7e7e9c418'
            WebId        = '1fe70582-a80c-470e-a8c2-8a9cf64e15a8'
            ListId       = 'A0049D75-96EA-48E2-BBA4-443D5C7D5175'
            FolderId     = '9050e1a2-39f3-4063-8afd-c64c5dc43851'
            WebUrl       = 'https://wstllc.sharepoint.com/sites/admin'
            FolderUrl    = 'https://wstllc.sharepoint.com/sites/admin/Shared Documents/General'
        },
        @{
            TenantTitle  = 'WST'
            WebTitle     = 'Admin/Accounting'
            ChannelTitle = 'General'
            GroupId      = '501d6636-c197-4307-8b01-556fc891fd89'
            SiteId       = '9c8e611d-7814-4d2d-9455-7e09b1c3f9d4'
            WebId        = '1fe70582-a80c-470e-a8c2-8a9cf64e15a8'
            ListId       = 'A0049D75-96EA-48E2-BBA4-443D5C7D5175'
            FolderId     = '67cd41da-c78d-4f84-8cc6-0a8d260412aa'
            WebUrl       = 'https://wstllc.sharepoint.com/sites/adminaccounting'
            FolderUrl    = 'https://wstllc.sharepoint.com/sites/adminaccounting/Shared Documents/General'
        }
    )}
}
#>

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body | ConvertTo-Json
})
