using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
}

Connect-PnPOnline -Url "https://edgetgpoc.sharepoint.com/sites/ExternalSharingTest" -Thumbprint "21dc042bb3127ebea6fc79c37e01bd86bde8d659" -ClientId "295ea9ee-f57a-46a5-877a-2c0620025369" -Tenant "edgepoc.co.uk"

$Site = Get-PnPSite
$SiteURL = $Site.URL
Write-Host $SiteURL -f Green

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
