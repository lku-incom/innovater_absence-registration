# Sync Absence Registrations to Holiday Balance
# Updates UsedDays/FeriefridageUsed based on approved absences

param(
    [switch]$WhatIf,         # Preview changes without applying
    [switch]$FullRecalculate # Recalculate all from scratch (ignores existing values)
)

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

Write-Host "=== Sync Absence to Holiday Balance ===" -ForegroundColor Magenta
if ($WhatIf) {
    Write-Host "[PREVIEW MODE - No changes will be made]" -ForegroundColor Yellow
}
if ($FullRecalculate) {
    Write-Host "[FULL RECALCULATE MODE - Will reset and recalculate all used days]" -ForegroundColor Yellow
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

# Function to determine holiday year from a date
function Get-HolidayYear {
    param([DateTime]$Date)
    if ($Date.Month -ge 9) {
        return "$($Date.Year)-$($Date.Year + 1)"
    } else {
        return "$($Date.Year - 1)-$($Date.Year)"
    }
}

# Get all holiday balances
Write-Host "`n=== Fetching Holiday Balances ===" -ForegroundColor Cyan
$balances = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=cr_isactive eq true" -Headers $headers
Write-Host "Found $($balances.value.Count) active balance(s)"

# Get all approved absences (using cr153_ prefix for this table)
Write-Host "`n=== Fetching Approved Absences ===" -ForegroundColor Cyan
$absenceFilter = "cr153_status eq 100000002"  # 100000002 = Godkendt
$absences = Invoke-RestMethod -Uri "$apiUrl/cr153_absenceregistrations?`$filter=$([System.Uri]::EscapeDataString($absenceFilter))" -Headers $headers
Write-Host "Found $($absences.value.Count) approved absence(s)"

# Get all pending absences
Write-Host "`n=== Fetching Pending Absences ===" -ForegroundColor Cyan
$pendingFilter = "cr153_status eq 100000001"  # 100000001 = Afventer godkendelse
$pendingAbsences = Invoke-RestMethod -Uri "$apiUrl/cr153_absenceregistrations?`$filter=$([System.Uri]::EscapeDataString($pendingFilter))" -Headers $headers
Write-Host "Found $($pendingAbsences.value.Count) pending absence(s)"

# Group absences by employee email and holiday year
$absencesByEmployee = @{}
$pendingByEmployee = @{}

# Absence type option set values
$ABSENCE_TYPE_FERIE = 100000000
$ABSENCE_TYPE_FERIEFRIDAGE = 100000003

foreach ($absence in $absences.value) {
    $email = $absence.cr153_employeeemail
    $startDate = [DateTime]$absence.cr153_startdate
    $holidayYear = Get-HolidayYear -Date $startDate
    $key = "$email|$holidayYear"

    if (-not $absencesByEmployee.ContainsKey($key)) {
        $absencesByEmployee[$key] = @{
            FeriedageUsed = 0
            FeriefridageUsed = 0
            Absences = @()
        }
    }

    $days = if ($null -ne $absence.cr153_numberofdays) { [decimal]$absence.cr153_numberofdays } else { 0 }
    $absenceType = $absence.cr153_absencetype

    if ($absenceType -eq $ABSENCE_TYPE_FERIEFRIDAGE) {
        $absencesByEmployee[$key].FeriefridageUsed += $days
    } elseif ($absenceType -eq $ABSENCE_TYPE_FERIE) {
        # Only count Ferie towards feriedage balance
        $absencesByEmployee[$key].FeriedageUsed += $days
    }
    # Note: Sygdom, Barsel, Flex, Andet are not counted against holiday balance
    $absencesByEmployee[$key].Absences += $absence
}

foreach ($absence in $pendingAbsences.value) {
    $email = $absence.cr153_employeeemail
    $startDate = [DateTime]$absence.cr153_startdate
    $holidayYear = Get-HolidayYear -Date $startDate
    $key = "$email|$holidayYear"

    if (-not $pendingByEmployee.ContainsKey($key)) {
        $pendingByEmployee[$key] = @{
            FeriedagePending = 0
            FeriefridagePending = 0
        }
    }

    $days = if ($null -ne $absence.cr153_numberofdays) { [decimal]$absence.cr153_numberofdays } else { 0 }
    $absenceType = $absence.cr153_absencetype

    if ($absenceType -eq $ABSENCE_TYPE_FERIEFRIDAGE) {
        $pendingByEmployee[$key].FeriefridagePending += $days
    } elseif ($absenceType -eq $ABSENCE_TYPE_FERIE) {
        $pendingByEmployee[$key].FeriedagePending += $days
    }
}

Write-Host "`n=== Processing Balances ===" -ForegroundColor Cyan

$updated = 0
$unchanged = 0
$errors = 0

foreach ($balance in $balances.value) {
    $email = $balance.cr_employeeemail
    $holidayYear = $balance.cr_holidayyear
    $key = "$email|$holidayYear"
    $employeeName = $balance.cr_employeename

    Write-Host "`n--- $employeeName ($holidayYear) ---" -ForegroundColor White

    # Get current values
    $currentUsedDays = if ($null -ne $balance.cr_useddays) { [decimal]$balance.cr_useddays } else { 0 }
    $currentFeriefridageUsed = if ($null -ne $balance.cr_feriefridageused) { [decimal]$balance.cr_feriefridageused } else { 0 }
    $currentPendingDays = if ($null -ne $balance.cr_pendingdays) { [decimal]$balance.cr_pendingdays } else { 0 }
    $accruedDays = if ($null -ne $balance.cr_accrueddays) { [decimal]$balance.cr_accrueddays } else { 0 }
    $transferredIn = if ($null -ne $balance.cr_transferredindays) { [decimal]$balance.cr_transferredindays } else { 0 }
    $feriefridageAccrued = if ($null -ne $balance.cr_feriefridageaccrued) { [decimal]$balance.cr_feriefridageaccrued } else { 0 }
    $feriefridageTransferredIn = if ($null -ne $balance.cr_feriefridagetransferredin) { [decimal]$balance.cr_feriefridagetransferredin } else { 0 }

    # Get calculated values from absences
    $calculatedUsed = 0
    $calculatedFeriefridageUsed = 0
    $calculatedPending = 0
    $calculatedFeriefridagePending = 0

    if ($absencesByEmployee.ContainsKey($key)) {
        $calculatedUsed = $absencesByEmployee[$key].FeriedageUsed
        $calculatedFeriefridageUsed = $absencesByEmployee[$key].FeriefridageUsed
    }

    if ($pendingByEmployee.ContainsKey($key)) {
        $calculatedPending = $pendingByEmployee[$key].FeriedagePending
        $calculatedFeriefridagePending = $pendingByEmployee[$key].FeriefridagePending
    }

    Write-Host "  Current:    UsedDays=$currentUsedDays, FeriefridageUsed=$currentFeriefridageUsed, Pending=$currentPendingDays"
    Write-Host "  Calculated: UsedDays=$calculatedUsed, FeriefridageUsed=$calculatedFeriefridageUsed, Pending=$calculatedPending"

    # Check if update is needed
    $needsUpdate = $false
    if ($FullRecalculate) {
        $needsUpdate = $true
    } elseif ($currentUsedDays -ne $calculatedUsed -or
              $currentFeriefridageUsed -ne $calculatedFeriefridageUsed -or
              $currentPendingDays -ne $calculatedPending) {
        $needsUpdate = $true
    }

    if (-not $needsUpdate) {
        Write-Host "  No changes needed" -ForegroundColor Gray
        $unchanged++
        continue
    }

    # Calculate new available days
    $newAvailableDays = [math]::Round($accruedDays + $transferredIn - $calculatedUsed - $calculatedPending, 2)
    $newFeriefridageAvailable = [math]::Round($feriefridageAccrued + $feriefridageTransferredIn - $calculatedFeriefridageUsed, 2)

    Write-Host "  New Available: Feriedage=$newAvailableDays, Feriefridage=$newFeriefridageAvailable" -ForegroundColor Cyan

    if ($WhatIf) {
        Write-Host "  [PREVIEW] Would update record" -ForegroundColor Yellow
        $updated++
        continue
    }

    # Update the balance
    try {
        $updateData = @{
            "cr_useddays" = $calculatedUsed
            "cr_feriefridageused" = $calculatedFeriefridageUsed
            "cr_pendingdays" = $calculatedPending
            "cr_availabledays" = $newAvailableDays
            "cr_feriefridageavailable" = $newFeriefridageAvailable
        }

        $updateHeaders = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version" = "4.0"
        }

        Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($($balance.cr_holidaybalanceid))" -Method PATCH -Headers $updateHeaders -Body ($updateData | ConvertTo-Json)
        Write-Host "  Updated successfully" -ForegroundColor Green
        $updated++
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        $errors++
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "  Updated: $updated"
Write-Host "  Unchanged: $unchanged"
Write-Host "  Errors: $errors"

if ($WhatIf) {
    Write-Host "`n[PREVIEW MODE] No changes were made." -ForegroundColor Yellow
    Write-Host "Run without -WhatIf to apply changes." -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Green
