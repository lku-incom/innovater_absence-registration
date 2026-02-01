# Explore-SharePointList.ps1
# Shows columns and sample data from the Ferieregistrering SharePoint list

param(
    [string]$SiteUrl = "https://innovaterdk.sharepoint.com/sites/projektstyring",
    [string]$ListName = "Ferieregistrering"
)

# Check if PnP PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}

Import-Module PnP.PowerShell

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "`n=== List Fields ===" -ForegroundColor Magenta
$fields = Get-PnPField -List $ListName | Where-Object { -not $_.Hidden } | Select-Object Title, InternalName, TypeAsString
$fields | Format-Table -AutoSize

Write-Host "`n=== Sample Data (first 5 items) ===" -ForegroundColor Magenta
$items = Get-PnPListItem -List $ListName -PageSize 5 | Select-Object -First 5
foreach ($item in $items) {
    Write-Host "`nItem ID: $($item.Id)" -ForegroundColor Green
    $item.FieldValues.GetEnumerator() | Where-Object { $_.Key -notlike "_*" -and $_.Key -notlike "OData*" } | ForEach-Object {
        Write-Host "  $($_.Key): $($_.Value)"
    }
}

Write-Host "`n=== Total Item Count ===" -ForegroundColor Magenta
$list = Get-PnPList -Identity $ListName
Write-Host "Total items: $($list.ItemCount)" -ForegroundColor Green

Disconnect-PnPOnline
