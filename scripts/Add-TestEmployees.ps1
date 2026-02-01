# Add Multiple Test Employees with Holiday Balances
# Creates realistic holiday balance records for testing the employee selector

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

# Test employees to create
$testEmployees = @(
    @{ Name = "Anders Jensen"; Email = "anders.jensen@test.dk" },
    @{ Name = "Mette Nielsen"; Email = "mette.nielsen@test.dk" },
    @{ Name = "Lars Pedersen"; Email = "lars.pedersen@test.dk" },
    @{ Name = "Sofie Andersen"; Email = "sofie.andersen@test.dk" }
)

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
    "Prefer" = "return=representation"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Calculate current holiday year (Sept 1 - Aug 31)
$now = Get-Date
$year = $now.Year
if ($now.Month -ge 9) {
    $holidayYear = "$year-$($year + 1)"
} else {
    $holidayYear = "$($year - 1)-$year"
}

Write-Host "`nHoliday Year: $holidayYear" -ForegroundColor Cyan
Write-Host "Creating test employee balances...`n" -ForegroundColor Yellow

$monthsIntoYear = if ($now.Month -ge 9) { $now.Month - 8 } else { $now.Month + 4 }

foreach ($employee in $testEmployees) {
    Write-Host "--- $($employee.Name) ($($employee.Email)) ---" -ForegroundColor Cyan

    # Check if record already exists
    $filter = "cr_employeeemail eq '$($employee.Email)' and cr_holidayyear eq '$holidayYear'"
    $existingRecords = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($filter))" -Headers $headers

    # Generate varied data for each employee
    $random = Get-Random -Minimum 0 -Maximum 100
    $usageRate = 0.2 + ($random / 250)  # 20-60% usage rate
    $accruedDays = [math]::Round($monthsIntoYear * 2.08, 2)
    $usedDays = [math]::Round($accruedDays * $usageRate, 1)
    $pendingDays = Get-Random -Minimum 0 -Maximum 4
    $transferredInDays = Get-Random -Minimum 0 -Maximum 5

    $feriefridageAccrued = [math]::Round($monthsIntoYear * 0.42, 2)
    $feriefridageUsed = [math]::Round($feriefridageAccrued * $usageRate * 0.8, 1)

    $availableDays = $accruedDays + $transferredInDays - $usedDays - $pendingDays
    $feriefridageAvailable = $feriefridageAccrued - $feriefridageUsed

    $balanceRecord = @{
        "cr_name" = "$($employee.Name) - $holidayYear"
        "cr_employeename" = $employee.Name
        "cr_employeeemail" = $employee.Email
        "cr_holidayyear" = $holidayYear
        "cr_accrueddays" = $accruedDays
        "cr_useddays" = $usedDays
        "cr_pendingdays" = $pendingDays
        "cr_availabledays" = $availableDays
        "cr_carriedoverdays" = 0
        "cr_transferredindays" = $transferredInDays
        "cr_transferredoutdays" = 0
        "cr_hastransferagreement" = ($transferredInDays -gt 0)
        "cr_feriefridageaccrued" = $feriefridageAccrued
        "cr_feriefridageused" = $feriefridageUsed
        "cr_feriefridageavailable" = $feriefridageAvailable
        "cr_feriefridagetransferredin" = 0
        "cr_feriefridagetransferredout" = 0
        "cr_isactive" = $true
    }

    try {
        if ($existingRecords.value.Count -gt 0) {
            $recordId = $existingRecords.value[0].cr_holidaybalanceid
            $updateHeaders = @{
                "Authorization" = "Bearer $token"
                "Content-Type" = "application/json"
                "OData-MaxVersion" = "4.0"
                "OData-Version" = "4.0"
            }
            Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($recordId)" -Method PATCH -Headers $updateHeaders -Body ($balanceRecord | ConvertTo-Json -Depth 10)
            Write-Host "  Updated existing record" -ForegroundColor Yellow
        } else {
            $response = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances" -Method POST -Headers $headers -Body ($balanceRecord | ConvertTo-Json -Depth 10)
            Write-Host "  Created new record" -ForegroundColor Green
        }
        Write-Host "  Feriedage: Accrued=$accruedDays, Used=$usedDays, Available=$availableDays"
        Write-Host "  Feriefridage: Accrued=$feriefridageAccrued, Used=$feriefridageUsed, Available=$feriefridageAvailable"
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
    }
    Write-Host ""
}

# Also fix the existing record with empty email (if any)
Write-Host "=== Checking for records with empty emails ===" -ForegroundColor Cyan
$emptyEmailFilter = "cr_employeeemail eq '' or cr_employeeemail eq null"
try {
    $emptyRecords = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($emptyEmailFilter))" -Headers $headers -ErrorAction SilentlyContinue
    if ($emptyRecords.value.Count -gt 0) {
        Write-Host "Found $($emptyRecords.value.Count) records with empty email. Deleting..." -ForegroundColor Yellow
        foreach ($record in $emptyRecords.value) {
            $deleteHeaders = @{
                "Authorization" = "Bearer $token"
                "OData-MaxVersion" = "4.0"
                "OData-Version" = "4.0"
            }
            Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($($record.cr_holidaybalanceid))" -Method DELETE -Headers $deleteHeaders
            Write-Host "  Deleted: $($record.cr_name)"
        }
    } else {
        Write-Host "No records with empty emails found." -ForegroundColor Green
    }
} catch {
    Write-Host "Could not check for empty email records: $_" -ForegroundColor Yellow
}

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "Created/updated $($testEmployees.Count) test employee balances for holiday year $holidayYear"
Write-Host "`nYou can now test the employee selector in the SPFx web part!"
Write-Host "Test employees:"
foreach ($employee in $testEmployees) {
    Write-Host "  - $($employee.Name) ($($employee.Email))"
}
