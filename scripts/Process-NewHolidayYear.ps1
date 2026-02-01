# Process New Holiday Year
# Creates new year balances and processes transfers from previous year

param(
    [switch]$WhatIf,     # Preview changes without applying
    [string]$ForYear     # Optional: Process specific year (format: "2025-2026")
)

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

Write-Host "=== New Holiday Year Setup ===" -ForegroundColor Magenta
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
    "Prefer" = "return=representation"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Determine years to process
$now = Get-Date
if ($ForYear) {
    $newYear = $ForYear
    $parts = $ForYear.Split('-')
    $previousYear = "$([int]$parts[0] - 1)-$($parts[0])"
} else {
    # Calculate based on current date
    if ($now.Month -ge 9) {
        $newYear = "$($now.Year)-$($now.Year + 1)"
        $previousYear = "$($now.Year - 1)-$($now.Year)"
    } else {
        $newYear = "$($now.Year - 1)-$($now.Year)"
        $previousYear = "$($now.Year - 2)-$($now.Year - 1)"
    }
}

Write-Host "`n=== Year Parameters ===" -ForegroundColor Cyan
Write-Host "  Previous Year: $previousYear"
Write-Host "  New Year: $newYear"

# Get balances from previous year
Write-Host "`n=== Fetching Previous Year Balances ===" -ForegroundColor Cyan
$filter = "cr_holidayyear eq '$previousYear'"
$previousBalances = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($filter))" -Headers $headers
Write-Host "Found $($previousBalances.value.Count) balance(s) for $previousYear"

# Check for existing new year balances
Write-Host "`n=== Checking Existing New Year Balances ===" -ForegroundColor Cyan
$filter = "cr_holidayyear eq '$newYear'"
$existingNewYear = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($filter))" -Headers $headers
Write-Host "Found $($existingNewYear.value.Count) existing balance(s) for $newYear"

# Create lookup for existing new year balances
$existingByEmail = @{}
foreach ($balance in $existingNewYear.value) {
    $existingByEmail[$balance.cr_employeeemail] = $balance
}

$created = 0
$updated = 0
$skipped = 0
$errors = 0
$transfersProcessed = 0

Write-Host "`n=== Processing Employees ===" -ForegroundColor Cyan

foreach ($prevBalance in $previousBalances.value) {
    $employeeName = $prevBalance.cr_employeename
    $employeeEmail = $prevBalance.cr_employeeemail

    Write-Host "`n--- $employeeName ---" -ForegroundColor White

    # Calculate transfer amounts
    $hasTransferAgreement = $prevBalance.cr_hastransferagreement -eq $true
    $transferredOutDays = if ($null -ne $prevBalance.cr_transferredoutdays) { [decimal]$prevBalance.cr_transferredoutdays } else { 0 }
    $feriefridageTransferredOut = if ($null -ne $prevBalance.cr_feriefridagetransferredout) { [decimal]$prevBalance.cr_feriefridagetransferredout } else { 0 }

    $transferInDays = 0
    $feriefridageTransferIn = 0

    if ($hasTransferAgreement -and $transferredOutDays -gt 0) {
        $transferInDays = $transferredOutDays
        $feriefridageTransferIn = $feriefridageTransferredOut
        Write-Host "  Transfer Agreement: YES - Transferring $transferInDays feriedage, $feriefridageTransferIn feriefridage" -ForegroundColor Green
        $transfersProcessed++
    } else {
        Write-Host "  Transfer Agreement: NO - No days transferred" -ForegroundColor Gray
    }

    # Check if new year balance already exists
    if ($existingByEmail.ContainsKey($employeeEmail)) {
        $existingBalance = $existingByEmail[$employeeEmail]
        Write-Host "  New year balance exists (ID: $($existingBalance.cr_holidaybalanceid))"

        # Check if we need to update transfer values
        $existingTransferIn = if ($null -ne $existingBalance.cr_transferredindays) { [decimal]$existingBalance.cr_transferredindays } else { 0 }

        if ($existingTransferIn -ne $transferInDays) {
            Write-Host "  Updating transfer: $existingTransferIn -> $transferInDays"

            if (-not $WhatIf) {
                try {
                    $updateData = @{
                        "cr_transferredindays" = $transferInDays
                        "cr_feriefridagetransferredin" = $feriefridageTransferIn
                        "cr_availabledays" = $transferInDays
                        "cr_feriefridageavailable" = $feriefridageTransferIn
                    }

                    $updateHeaders = @{
                        "Authorization" = "Bearer $token"
                        "Content-Type" = "application/json"
                        "OData-MaxVersion" = "4.0"
                        "OData-Version" = "4.0"
                    }

                    Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($($existingBalance.cr_holidaybalanceid))" -Method PATCH -Headers $updateHeaders -Body ($updateData | ConvertTo-Json)
                    Write-Host "  Updated" -ForegroundColor Green
                    $updated++
                } catch {
                    Write-Host "  ERROR: $_" -ForegroundColor Red
                    $errors++
                }
            } else {
                Write-Host "  [PREVIEW] Would update" -ForegroundColor Yellow
                $updated++
            }
        } else {
            Write-Host "  No update needed" -ForegroundColor Gray
            $skipped++
        }
        continue
    }

    # Create new year balance
    Write-Host "  Creating new balance for $newYear"

    $newBalance = @{
        "cr_name" = "$employeeName - $newYear"
        "cr_employeeemail" = $employeeEmail
        "cr_employeename" = $employeeName
        "cr_holidayyear" = $newYear
        "cr_accrueddays" = 0
        "cr_useddays" = 0
        "cr_pendingdays" = 0
        "cr_availabledays" = $transferInDays
        "cr_carriedoverdays" = 0
        "cr_transferredindays" = $transferInDays
        "cr_transferredoutdays" = 0
        "cr_hastransferagreement" = $false
        "cr_feriefridageaccrued" = 0
        "cr_feriefridageused" = 0
        "cr_feriefridageavailable" = $feriefridageTransferIn
        "cr_feriefridagetransferredin" = $feriefridageTransferIn
        "cr_feriefridagetransferredout" = 0
        "cr_isactive" = $true
    }

    if ($WhatIf) {
        Write-Host "  [PREVIEW] Would create new balance" -ForegroundColor Yellow
        $created++
        continue
    }

    try {
        $response = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances" -Method POST -Headers $headers -Body ($newBalance | ConvertTo-Json -Depth 10)
        Write-Host "  Created (ID: $($response.cr_holidaybalanceid))" -ForegroundColor Green
        $created++
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        $errors++
    }
}

# Mark previous year as inactive
Write-Host "`n=== Marking Previous Year as Inactive ===" -ForegroundColor Cyan

foreach ($prevBalance in $previousBalances.value) {
    if ($prevBalance.cr_isactive -eq $true) {
        if ($WhatIf) {
            Write-Host "  [PREVIEW] Would mark $($prevBalance.cr_employeename) as inactive" -ForegroundColor Yellow
            continue
        }

        try {
            $updateHeaders = @{
                "Authorization" = "Bearer $token"
                "Content-Type" = "application/json"
                "OData-MaxVersion" = "4.0"
                "OData-Version" = "4.0"
            }

            Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($($prevBalance.cr_holidaybalanceid))" -Method PATCH -Headers $updateHeaders -Body '{"cr_isactive": false}'
            Write-Host "  Marked $($prevBalance.cr_employeename) as inactive" -ForegroundColor Gray
        } catch {
            Write-Host "  ERROR marking inactive: $_" -ForegroundColor Red
        }
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "  New Year: $newYear"
Write-Host "  Created: $created"
Write-Host "  Updated: $updated"
Write-Host "  Skipped: $skipped"
Write-Host "  Errors: $errors"
Write-Host "  Transfers Processed: $transfersProcessed"

if ($WhatIf) {
    Write-Host "`n[PREVIEW MODE] No changes were made." -ForegroundColor Yellow
    Write-Host "Run without -WhatIf to apply changes." -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Green
