# Test Holiday Transfer Logic
# Demonstrates the Danish holiday law transfer rules

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

# Authenticate
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
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Fetch a sample employee's balance
Write-Host "`n=== Fetching Sample Holiday Balance ===" -ForegroundColor Cyan
$balances = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$top=3&`$orderby=cr_employeename" -Headers $headers

Write-Host "`nFound $($balances.value.Count) balance records. Testing with first one:`n"

$testBalance = $balances.value[0]
Write-Host "Employee: $($testBalance.cr_employeename)" -ForegroundColor Yellow
Write-Host "Holiday Year: $($testBalance.cr_holidayyear)"
Write-Host "Accrued Days: $($testBalance.cr_accrueddays)"
Write-Host "Used Days: $($testBalance.cr_useddays)"
Write-Host "Pending Days: $($testBalance.cr_pendingdays)"
Write-Host "Carried Over (legacy): $($testBalance.cr_carriedoverdays)"
Write-Host "Transferred In: $($testBalance.cr_transferredindays)"
Write-Host "Transferred Out: $($testBalance.cr_transferredoutdays)"
Write-Host "Has Transfer Agreement: $($testBalance.cr_hastransferagreement)"
Write-Host "Transfer Agreement Date: $($testBalance.cr_transferagreementdate)"

# Calculate transfer eligibility (Danish Holiday Law logic)
Write-Host "`n=== Transfer Eligibility Analysis ===" -ForegroundColor Magenta

$MANDATORY_DAYS = 20
$MAX_TRANSFER = 5

$accrued = if ($null -ne $testBalance.cr_accrueddays) { [decimal]$testBalance.cr_accrueddays } else { 0 }
$used = if ($null -ne $testBalance.cr_useddays) { [decimal]$testBalance.cr_useddays } else { 0 }
$pending = if ($null -ne $testBalance.cr_pendingdays) { [decimal]$testBalance.cr_pendingdays } else { 0 }
$transferredIn = if ($null -ne $testBalance.cr_transferredindays) { [decimal]$testBalance.cr_transferredindays } else { 0 }
$carriedOver = if ($null -ne $testBalance.cr_carriedoverdays) { [decimal]$testBalance.cr_carriedoverdays } else { 0 }

$totalAvailable = $accrued + $transferredIn + $carriedOver
$remaining = $totalAvailable - $used - $pending
$daysAboveMandatory = [Math]::Max(0, $remaining - $MANDATORY_DAYS)
$maxTransferable = [Math]::Min($daysAboveMandatory, $MAX_TRANSFER)

Write-Host "`nCalculations:"
Write-Host "  Total Available: $totalAvailable days (accrued: $accrued + transferred in: $transferredIn + carried over: $carriedOver)"
Write-Host "  Remaining after usage: $remaining days (available - used: $used - pending: $pending)"
Write-Host "  Days above mandatory 20: $daysAboveMandatory"
Write-Host "  Max transferable (capped at 5): $maxTransferable" -ForegroundColor Green

# Check mandatory vacation status
$mandatoryTaken = $used -ge $MANDATORY_DAYS
$daysAtRisk = [Math]::Min([Math]::Max(0, $MANDATORY_DAYS - $used), $remaining)

Write-Host "`nMandatory Vacation Status:"
if ($mandatoryTaken) {
    Write-Host "  Mandatory 20 days: TAKEN" -ForegroundColor Green
} else {
    Write-Host "  Mandatory 20 days: NOT YET MET ($used of 20 taken)" -ForegroundColor Yellow
    Write-Host "  Days at risk of forfeiture: $daysAtRisk" -ForegroundColor Red
}

# Simulate a transfer
Write-Host "`n=== Simulating Transfer Request ===" -ForegroundColor Cyan

if ($maxTransferable -gt 0) {
    $daysToTransfer = [Math]::Min(3, $maxTransferable)  # Transfer 3 days or max available

    Write-Host "Simulating transfer of $daysToTransfer days..."

    # Parse holiday year to get next year
    $yearParts = $testBalance.cr_holidayyear -split '-'
    $nextHolidayYear = "$($yearParts[1])-$([int]$yearParts[1] + 1)"

    Write-Host "`nTransfer would:"
    Write-Host "  1. Set TransferredOutDays = $daysToTransfer on current balance"
    Write-Host "  2. Set HasTransferAgreement = true"
    Write-Host "  3. Set TransferAgreementDate = $(Get-Date -Format 'yyyy-MM-dd')"
    Write-Host "  4. Create/update $nextHolidayYear balance with TransferredInDays = $daysToTransfer"

    # Ask if user wants to actually perform the transfer
    Write-Host "`n=== Test Update ===" -ForegroundColor Magenta
    Write-Host "Would you like to test updating the transfer fields? (This will update the record)"
    $confirm = Read-Host "Type 'yes' to proceed, anything else to skip"

    if ($confirm -eq 'yes') {
        $updateBody = @{
            "cr_transferredoutdays" = $daysToTransfer
            "cr_hastransferagreement" = $true
            "cr_transferagreementdate" = (Get-Date -Format 'yyyy-MM-dd')
        } | ConvertTo-Json

        try {
            $recordId = $testBalance.cr_holidaybalanceid
            Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($recordId)" -Method PATCH -Headers $headers -Body $updateBody
            Write-Host "SUCCESS! Updated transfer fields for $($testBalance.cr_employeename)" -ForegroundColor Green

            # Verify the update
            $updated = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($recordId)" -Headers $headers
            Write-Host "`nVerified values:"
            Write-Host "  Transferred Out Days: $($updated.cr_transferredoutdays)"
            Write-Host "  Has Transfer Agreement: $($updated.cr_hastransferagreement)"
            Write-Host "  Transfer Agreement Date: $($updated.cr_transferagreementdate)"
        }
        catch {
            Write-Host "ERROR: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "Skipped actual update." -ForegroundColor Yellow
    }
} else {
    Write-Host "No days available for transfer. Employee must take more vacation first." -ForegroundColor Yellow
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Green
