# Truncate All Tables
# Deletes all records from AbsenceRegistration, HolidayBalance, and AccrualHistory tables

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

Write-Host "=== Truncate All Tables ===" -ForegroundColor Magenta
Write-Host "This will DELETE ALL records from:" -ForegroundColor Red
Write-Host "  - cr153_absenceregistrations" -ForegroundColor Yellow
Write-Host "  - cr_holidaybalances" -ForegroundColor Yellow
Write-Host "  - cr_accrualhistories" -ForegroundColor Yellow
Write-Host ""

$confirm = Read-Host "Are you sure? Type 'YES' to confirm"
if ($confirm -ne "YES") {
    Write-Host "Cancelled." -ForegroundColor Yellow
    exit
}

# Authenticate
Write-Host "`n=== Authenticating ===" -ForegroundColor Cyan
$deviceCodeUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode"
$tokenUrl = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
$scope = "$DataverseUrl/.default"

$deviceCodeResponse = Invoke-RestMethod -Uri $deviceCodeUrl -Method POST -Body @{
    client_id = $clientId
    scope = $scope
}

Write-Host $deviceCodeResponse.message -ForegroundColor Yellow

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
        if ($_.Exception.Response.StatusCode -ne 400) { throw }
    }
}

$headers = @{
    "Authorization" = "Bearer $token"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Function to delete all records from a table
function Clear-Table {
    param(
        [string]$TableName,
        [string]$IdField
    )

    Write-Host "`n=== Truncating $TableName ===" -ForegroundColor Cyan

    $allRecords = @()
    $nextLink = "$apiUrl/$TableName"

    # Get all records (handle pagination)
    while ($nextLink) {
        $response = Invoke-RestMethod -Uri $nextLink -Headers $headers
        $allRecords += $response.value
        $nextLink = $response.'@odata.nextLink'
    }

    Write-Host "Found $($allRecords.Count) records to delete"

    if ($allRecords.Count -eq 0) {
        Write-Host "Table is already empty." -ForegroundColor Gray
        return 0
    }

    $deleted = 0
    $errors = 0

    foreach ($record in $allRecords) {
        $recordId = $record.$IdField
        try {
            Invoke-RestMethod -Uri "$apiUrl/$TableName($recordId)" -Method DELETE -Headers $headers
            $deleted++
            Write-Host "." -NoNewline
        } catch {
            $errors++
            Write-Host "x" -NoNewline -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "Deleted: $deleted, Errors: $errors" -ForegroundColor $(if ($errors -eq 0) { "Green" } else { "Yellow" })

    return $deleted
}

$totalDeleted = 0

# 1. Truncate Absence Registrations
$totalDeleted += Clear-Table -TableName "cr153_absenceregistrations" -IdField "cr153_absenceregistrationid"

# 2. Truncate Holiday Balances
$totalDeleted += Clear-Table -TableName "cr_holidaybalances" -IdField "cr_holidaybalanceid"

# 3. Truncate Accrual History
$totalDeleted += Clear-Table -TableName "cr_accrualhistories" -IdField "cr_accrualhistoryid"

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "Total records deleted: $totalDeleted" -ForegroundColor Green
Write-Host "`nAll tables have been truncated." -ForegroundColor Green
