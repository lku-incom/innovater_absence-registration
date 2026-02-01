# Run Monthly Accrual - Manual Script
# Simulates what the Power Automate flow would do for monthly accrual

param(
    [switch]$WhatIf,  # Preview changes without applying
    [string]$ForMonth  # Optional: Run for specific month (format: "2025-02")
)

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

# Danish Holiday Law constants
$MONTHLY_ACCRUAL_RATE = 2.08       # Feriedage per month
$MONTHLY_FERIEFRIDAGE_RATE = 0.42 # Feriefridage per month (5/12)

Write-Host "=== Monthly Accrual Script ===" -ForegroundColor Magenta
if ($WhatIf) {
    Write-Host "[PREVIEW MODE - No changes will be made]" -ForegroundColor Yellow
}
Write-Host ""

# Authenticate
Write-Host "=== Authenticating ===" -ForegroundColor Cyan
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

# Determine processing date
if ($ForMonth) {
    $processingDate = [DateTime]::ParseExact("$ForMonth-01", "yyyy-MM-dd", $null)
} else {
    $processingDate = Get-Date
}

$accrualMonth = $processingDate.Month
$accrualYear = $processingDate.Year

# Calculate holiday year
if ($processingDate.Month -ge 9) {
    $holidayYear = "$($processingDate.Year)-$($processingDate.Year + 1)"
} else {
    $holidayYear = "$($processingDate.Year - 1)-$($processingDate.Year)"
}

Write-Host "`n=== Accrual Parameters ===" -ForegroundColor Cyan
Write-Host "  Processing Date: $($processingDate.ToString('yyyy-MM-dd'))"
Write-Host "  Holiday Year: $holidayYear"
Write-Host "  Accrual Month: $accrualMonth"
Write-Host "  Accrual Year: $accrualYear"
Write-Host "  Feriedage: +$MONTHLY_ACCRUAL_RATE days"
Write-Host "  Feriefridage: +$MONTHLY_FERIEFRIDAGE_RATE days"

# Get all active balances for the holiday year
Write-Host "`n=== Fetching Active Balances ===" -ForegroundColor Cyan
$filter = "cr_holidayyear eq '$holidayYear' and cr_isactive eq true"
$balances = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($filter))" -Headers $headers

Write-Host "Found $($balances.value.Count) active balance(s) for $holidayYear"

if ($balances.value.Count -eq 0) {
    Write-Host "`nNo balances to process. Exiting." -ForegroundColor Yellow
    exit
}

$processed = 0
$skipped = 0
$errors = 0

Write-Host "`n=== Processing Accruals ===" -ForegroundColor Cyan

foreach ($balance in $balances.value) {
    $employeeName = $balance.cr_employeename
    $employeeEmail = $balance.cr_employeeemail
    $currentAccrued = if ($null -ne $balance.cr_accrueddays) { [decimal]$balance.cr_accrueddays } else { 0 }
    $currentFeriefridage = if ($null -ne $balance.cr_feriefridageaccrued) { [decimal]$balance.cr_feriefridageaccrued } else { 0 }
    $lastAccrualDate = $balance.cr_lastaccrualdate

    Write-Host "`n--- $employeeName ---" -ForegroundColor White

    # Check if already accrued this month
    if ($lastAccrualDate) {
        $lastAccrual = [DateTime]$lastAccrualDate
        if ($lastAccrual.Year -eq $accrualYear -and $lastAccrual.Month -eq $accrualMonth) {
            Write-Host "  SKIPPED: Already accrued for $($lastAccrual.ToString('MMMM yyyy'))" -ForegroundColor Yellow
            $skipped++
            continue
        }
    }

    # Calculate new values
    $newAccruedDays = [math]::Round($currentAccrued + $MONTHLY_ACCRUAL_RATE, 2)
    $newFeriefridage = [math]::Round($currentFeriefridage + $MONTHLY_FERIEFRIDAGE_RATE, 2)

    # Recalculate available days
    $usedDays = if ($null -ne $balance.cr_useddays) { [decimal]$balance.cr_useddays } else { 0 }
    $pendingDays = if ($null -ne $balance.cr_pendingdays) { [decimal]$balance.cr_pendingdays } else { 0 }
    $transferredIn = if ($null -ne $balance.cr_transferredindays) { [decimal]$balance.cr_transferredindays } else { 0 }
    $feriefridageUsed = if ($null -ne $balance.cr_feriefridageused) { [decimal]$balance.cr_feriefridageused } else { 0 }
    $feriefridageTransferredIn = if ($null -ne $balance.cr_feriefridagetransferredin) { [decimal]$balance.cr_feriefridagetransferredin } else { 0 }

    $newAvailableDays = [math]::Round($newAccruedDays + $transferredIn - $usedDays - $pendingDays, 2)
    $newFeriefridageAvailable = [math]::Round($newFeriefridage + $feriefridageTransferredIn - $feriefridageUsed, 2)

    Write-Host "  Current: Accrued=$currentAccrued, Feriefridage=$currentFeriefridage"
    Write-Host "  New:     Accrued=$newAccruedDays (+$MONTHLY_ACCRUAL_RATE), Feriefridage=$newFeriefridage (+$MONTHLY_FERIEFRIDAGE_RATE)"
    Write-Host "  Available: Feriedage=$newAvailableDays, Feriefridage=$newFeriefridageAvailable"

    if ($WhatIf) {
        Write-Host "  [PREVIEW] Would update record" -ForegroundColor Cyan
        $processed++
        continue
    }

    # Update the balance record
    try {
        $updateData = @{
            "cr_accrueddays" = $newAccruedDays
            "cr_feriefridageaccrued" = $newFeriefridage
            "cr_availabledays" = $newAvailableDays
            "cr_feriefridageavailable" = $newFeriefridageAvailable
            "cr_lastaccrualdate" = $processingDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }

        $updateHeaders = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version" = "4.0"
        }

        Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($($balance.cr_holidaybalanceid))" -Method PATCH -Headers $updateHeaders -Body ($updateData | ConvertTo-Json)
        Write-Host "  Updated successfully" -ForegroundColor Green
        $processed++
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        $errors++
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "  Processed: $processed"
Write-Host "  Skipped (already accrued): $skipped"
Write-Host "  Errors: $errors"

if ($WhatIf) {
    Write-Host "`n[PREVIEW MODE] No changes were made." -ForegroundColor Yellow
    Write-Host "Run without -WhatIf to apply changes." -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Green
