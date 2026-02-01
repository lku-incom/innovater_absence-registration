# Quick check of holiday balance records
$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

Write-Host "=== Authenticating ===" -ForegroundColor Magenta
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

Write-Host "`n=== All Holiday Balance Records ===" -ForegroundColor Cyan
$response = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$select=cr_name,cr_employeename,cr_employeeemail,cr_holidayyear,cr_accrueddays,cr_useddays&`$orderby=cr_employeename" -Headers $headers

Write-Host "Total records: $($response.value.Count)`n"

foreach ($record in $response.value) {
    $email = if ($record.cr_employeeemail) { $record.cr_employeeemail } else { "(NO EMAIL)" }
    Write-Host "  $($record.cr_employeename) | $email | Year: $($record.cr_holidayyear) | Accrued: $($record.cr_accrueddays)"
}

Write-Host "`n=== Records for 2025-2026 with valid email ===" -ForegroundColor Cyan
$filtered = $response.value | Where-Object { $_.cr_holidayyear -eq "2025-2026" -and $_.cr_employeeemail }
Write-Host "Count: $($filtered.Count)"
foreach ($record in $filtered) {
    Write-Host "  $($record.cr_employeename) | $($record.cr_employeeemail)"
}

Write-Host "`n=== Done ===" -ForegroundColor Green
