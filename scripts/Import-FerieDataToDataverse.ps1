# Import-FerieDataToDataverse.ps1
# Imports historical Ferieregistrering data into Dataverse tables
# - Accrual records → cr_accrualhistory
# - Calculates balances → cr_holidaybalance

param(
    [string]$CsvPath = "c:\Users\info\Innovater\absence-registration\scripts\Ferieregistrering.csv",
    [string]$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com",
    [int]$StartYear = 2025  # Only import data from this year onwards
)

# Danish holiday law constants
$MONTHLY_ACCRUAL_RATE = 2.08
$HOLIDAY_YEAR_START_MONTH = 9  # September

function Get-HolidayYear {
    param([DateTime]$Date)
    $month = $Date.Month
    $year = $Date.Year
    if ($month -ge $HOLIDAY_YEAR_START_MONTH) {
        return "$year-$($year + 1)"
    } else {
        return "$($year - 1)-$year"
    }
}

function Get-AccrualType {
    param([string]$Category)
    switch ($Category) {
        "Tilskrivning, feriedage" { return 100000000 }  # Monthly Accrual
        "Tilskrivning, feriefridage" { return 100000000 }  # Monthly Accrual
        "Ferie" { return 100000002 }  # Manual Adjustment (negative = used)
        "Feriefridage" { return 100000002 }  # Manual Adjustment
        default { return 100000002 }  # Manual Adjustment
    }
}

# Authenticate to Dataverse
Write-Host "=== Authenticating to Dataverse ===" -ForegroundColor Magenta
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"
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

if (-not $token) {
    Write-Error "Failed to authenticate"
    exit 1
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
    "Prefer" = "return=representation"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Load CSV data
Write-Host "`n=== Loading CSV Data ===" -ForegroundColor Magenta
$allData = Import-Csv $CsvPath -Encoding UTF8
Write-Host "Loaded $($allData.Count) total records"

# Filter to only include data from StartYear onwards
Write-Host "Filtering to records from $StartYear onwards..." -ForegroundColor Yellow
$data = $allData | Where-Object {
    try {
        $date = [DateTime]::ParseExact($_.'Dato start', 'dd-MM-yyyy', $null)
        return $date.Year -ge $StartYear
    } catch {
        return $false
    }
}
Write-Host "Filtered to $($data.Count) records" -ForegroundColor Green

# Process and import data
Write-Host "`n=== De-duplicating Records ===" -ForegroundColor Magenta

$accrualCategories = @("Tilskrivning, feriedage", "Tilskrivning, feriefridage")
$usageCategories = @("Ferie", "Feriefridage")

# De-duplicate: Group by Employee + Date + Category, keep first record (duplicates are data entry errors)
$deduplicatedData = $data | Group-Object { "$($_.Medarbejder)|$($_.'Dato start')|$($_.Kategori)" } | ForEach-Object {
    if ($_.Count -gt 1) {
        Write-Host "  Duplicate found: $($_.Name) - keeping 1 of $($_.Count) records" -ForegroundColor Yellow
    }
    $_.Group | Select-Object -First 1
}
Write-Host "After de-duplication: $($deduplicatedData.Count) records (removed $($data.Count - $deduplicatedData.Count) duplicates)" -ForegroundColor Green

Write-Host "`n=== Importing Accrual History Records ===" -ForegroundColor Magenta

$successCount = 0
$errorCount = 0
$skipCount = 0

# Track balances per employee per holiday year
$balances = @{}

foreach ($row in $deduplicatedData) {
    $category = $row.Kategori

    # Parse the date
    try {
        $accrualDate = [DateTime]::ParseExact($row.'Dato start', 'dd-MM-yyyy', $null)
    } catch {
        Write-Warning "Could not parse date: $($row.'Dato start') for $($row.Medarbejder)"
        $errorCount++
        continue
    }

    # Parse the number (Danish format uses comma as decimal separator)
    $daysStr = $row.'Antal dage' -replace ',', '.'
    $days = [decimal]$daysStr

    $employeeName = $row.Medarbejder
    $holidayYear = Get-HolidayYear -Date $accrualDate
    $accrualMonth = $accrualDate.Month
    $accrualYear = $accrualDate.Year

    # Initialize balance tracking for this employee/year
    $balanceKey = "$employeeName|$holidayYear"
    if (-not $balances.ContainsKey($balanceKey)) {
        $balances[$balanceKey] = @{
            EmployeeName = $employeeName
            HolidayYear = $holidayYear
            AccruedDays = 0
            UsedDays = 0
            FeriefridageAccrued = 0
            FeriefridageUsed = 0
        }
    }

    # Only import accrual records to history table
    if ($category -in $accrualCategories) {
        # Determine if this is ferie or feriefridage accrual
        $isFeriefridage = $category -eq "Tilskrivning, feriefridage"

        # Update balance tracking
        if ($isFeriefridage) {
            $balances[$balanceKey].FeriefridageAccrued += $days
        } else {
            $balances[$balanceKey].AccruedDays += $days
        }

        # Create accrual history record
        # Include type suffix to ensure unique names (important for January with both feriedage + feriefridage)
        $typeSuffix = if ($isFeriefridage) { "Feriefridage" } else { "Feriedage" }
        $recordName = "$employeeName - $(Get-Date $accrualDate -Format 'MMM yyyy') - $typeSuffix"

        $accrualRecord = @{
            "cr_name" = $recordName
            "cr_employeename" = $employeeName
            "cr_employeeemail" = ""  # Not available in source data
            "cr_holidayyear" = $holidayYear
            "cr_accrualdate" = $accrualDate.ToString("yyyy-MM-dd")
            "cr_accrualmonth" = $accrualMonth
            "cr_accrualyear" = $accrualYear
            "cr_daysaccrued" = if ($isFeriefridage) { 0 } else { $days }
            "cr_feriefridageaccrued" = if ($isFeriefridage) { $days } else { $null }
            "cr_accrualtype" = Get-AccrualType -Category $category
            "cr_notes" = "Imported from SharePoint: $category"
        }

        try {
            $response = Invoke-RestMethod -Uri "$apiUrl/cr_accrualhistories" -Method POST -Headers $headers -Body ($accrualRecord | ConvertTo-Json -Depth 10)
            $successCount++
            if ($successCount % 50 -eq 0) {
                Write-Host "  Imported $successCount accrual records..." -ForegroundColor Gray
            }
        } catch {
            $errorCount++
            if ($errorCount -le 5) {
                Write-Warning "Error importing record for $employeeName : $_"
            }
        }
    }
    elseif ($category -in $usageCategories) {
        # Track usage (negative days mean vacation taken)
        $usedDays = [Math]::Abs($days)
        if ($category -eq "Feriefridage") {
            $balances[$balanceKey].FeriefridageUsed += $usedDays
        } else {
            $balances[$balanceKey].UsedDays += $usedDays
        }
        $skipCount++  # Not importing usage to accrual history
    }
    else {
        # Other categories (Sygdom, Andet fravær, etc.)
        $skipCount++
    }
}

Write-Host "`nAccrual History Import Complete:" -ForegroundColor Green
Write-Host "  Success: $successCount"
Write-Host "  Errors: $errorCount"
Write-Host "  Skipped (usage/other): $skipCount"

# Create Holiday Balance records
Write-Host "`n=== Creating Holiday Balance Records ===" -ForegroundColor Magenta

$balanceSuccess = 0
$balanceError = 0

foreach ($key in $balances.Keys) {
    $balance = $balances[$key]

    # Calculate available days
    $availableDays = $balance.AccruedDays - $balance.UsedDays
    $availableFeriefridage = $balance.FeriefridageAccrued - $balance.FeriefridageUsed

    # Determine if this is the current holiday year
    $currentHolidayYear = Get-HolidayYear -Date (Get-Date)
    $isActive = $balance.HolidayYear -eq $currentHolidayYear

    $balanceRecord = @{
        "cr_name" = "$($balance.EmployeeName) - $($balance.HolidayYear)"
        "cr_employeename" = $balance.EmployeeName
        "cr_employeeemail" = ""  # Not available
        "cr_holidayyear" = $balance.HolidayYear
        "cr_accrueddays" = $balance.AccruedDays
        "cr_useddays" = $balance.UsedDays
        "cr_pendingdays" = 0
        "cr_availabledays" = $availableDays
        "cr_carriedoverdays" = 0
        "cr_feriefridageaccrued" = $balance.FeriefridageAccrued
        "cr_feriefridageused" = $balance.FeriefridageUsed
        "cr_feriefridageavailable" = $availableFeriefridage
        "cr_isactive" = $isActive
    }

    try {
        $response = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances" -Method POST -Headers $headers -Body ($balanceRecord | ConvertTo-Json -Depth 10)
        $balanceSuccess++
    } catch {
        $balanceError++
        if ($balanceError -le 5) {
            Write-Warning "Error creating balance for $($balance.EmployeeName) $($balance.HolidayYear): $_"
        }
    }
}

Write-Host "`nHoliday Balance Import Complete:" -ForegroundColor Green
Write-Host "  Success: $balanceSuccess"
Write-Host "  Errors: $balanceError"

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "Migration Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Magenta
Write-Host "`nSummary:"
Write-Host "  - Accrual History records: $successCount"
Write-Host "  - Holiday Balance records: $balanceSuccess"
Write-Host "  - Unique employees: $($balances.Keys.Count)"
