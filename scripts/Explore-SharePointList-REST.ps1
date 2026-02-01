# Explore-SharePointList-REST.ps1
# Uses SharePoint REST API to explore the Ferieregistrering list

param(
    [string]$SiteUrl = "https://innovaterdk.sharepoint.com/sites/projektstyring",
    [string]$ListName = "Ferieregistrering"
)

$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d" # Power Platform CLI client ID

Write-Host "Authenticating to SharePoint..." -ForegroundColor Cyan

# Device code flow for authentication
$deviceCodeUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode"
$tokenUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
$scope = "https://innovaterdk.sharepoint.com/.default"

$deviceCodeResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -Body @{
    client_id = $clientId
    scope = $scope
}

Write-Host $deviceCodeResponse.message -ForegroundColor Yellow

# Poll for token
$pollInterval = $deviceCodeResponse.interval
$expiresIn = $deviceCodeResponse.expires_in
$startTime = Get-Date
$token = $null

while ((Get-Date) -lt $startTime.AddSeconds($expiresIn)) {
    Start-Sleep -Seconds $pollInterval
    try {
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body @{
            grant_type = "urn:ietf:params:oauth:grant-type:device_code"
            client_id = $clientId
            device_code = $deviceCodeResponse.device_code
        }
        $token = $tokenResponse.access_token
        Write-Host "Authentication successful!" -ForegroundColor Green
        break
    } catch {
        if ($_.Exception.Response.StatusCode -ne 400) {
            throw
        }
    }
}

if (-not $token) {
    Write-Error "Failed to obtain access token"
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept" = "application/json;odata=verbose"
}

# Get list fields
Write-Host "`n=== List Fields ===" -ForegroundColor Magenta
$fieldsUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/fields?`$filter=Hidden eq false"
try {
    $fieldsResponse = Invoke-RestMethod -Uri $fieldsUrl -Headers $headers -Method GET
    $fields = $fieldsResponse.d.results | Select-Object Title, InternalName, TypeAsString
    $fields | Format-Table -AutoSize
} catch {
    Write-Warning "Error getting fields: $_"
}

# Get sample items
Write-Host "`n=== Sample Data (first 10 items) ===" -ForegroundColor Magenta
$itemsUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items?`$top=10"
try {
    $itemsResponse = Invoke-RestMethod -Uri $itemsUrl -Headers $headers -Method GET
    $items = $itemsResponse.d.results

    if ($items.Count -gt 0) {
        Write-Host "Found $($items.Count) items" -ForegroundColor Green

        # Show first item's properties
        Write-Host "`nFirst item properties:" -ForegroundColor Cyan
        $firstItem = $items[0]
        $firstItem.PSObject.Properties | Where-Object { $_.Name -notlike "__*" -and $_.Name -notlike "*deferred*" } | ForEach-Object {
            if ($_.Value -ne $null -and $_.Value -ne "") {
                Write-Host "  $($_.Name): $($_.Value)"
            }
        }

        # Export to CSV for inspection
        $csvPath = "c:\Users\info\Innovater\absence-registration\scripts\ferieregistrering_sample.csv"
        $items | Select-Object * -ExcludeProperty __metadata, FirstUniqueAncestorSecurableObject, RoleAssignments, AttachmentFiles, ContentType, GetDlpPolicyTip, FieldValuesAsHtml, FieldValuesAsText, FieldValuesForEdit, File, Folder, LikedByInformation, ParentList, Properties, Versions | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nSample data exported to: $csvPath" -ForegroundColor Green
    } else {
        Write-Host "No items found in the list" -ForegroundColor Yellow
    }
} catch {
    Write-Warning "Error getting items: $_"
    Write-Host "Full error: $($_.Exception.Message)"
}

# Get list item count
Write-Host "`n=== List Info ===" -ForegroundColor Magenta
$listUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')"
try {
    $listResponse = Invoke-RestMethod -Uri $listUrl -Headers $headers -Method GET
    Write-Host "List Title: $($listResponse.d.Title)"
    Write-Host "Item Count: $($listResponse.d.ItemCount)"
} catch {
    Write-Warning "Error getting list info: $_"
}
