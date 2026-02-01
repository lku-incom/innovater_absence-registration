# Add Dummy Holiday Balance Data
# Creates a realistic holiday balance record for testing

param(
    [string]$EmployeeEmail = "",
    [string]$EmployeeName = ""
)

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
    "Prefer" = "return=representation"
}

$apiUrl = "$DataverseUrl/api/data/v9.2"

# Get current user info if not provided
if (-not $EmployeeEmail) {
    Write-Host "`n=== Getting Current User Info ===" -ForegroundColor Cyan
    try {
        $meResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers @{
            "Authorization" = "Bearer $token"
        } -ErrorAction SilentlyContinue
        $EmployeeEmail = $meResponse.mail
        $EmployeeName = $meResponse.displayName
    } catch {
        # If Graph fails, prompt for email
        $EmployeeEmail = Read-Host "Enter your email address"
        $EmployeeName = Read-Host "Enter your name"
    }
}

if (-not $EmployeeEmail) {
    $EmployeeEmail = Read-Host "Enter your email address"
}
if (-not $EmployeeName) {
    $EmployeeName = Read-Host "Enter your name"
}

Write-Host "`nCreating holiday balance for: $EmployeeName ($EmployeeEmail)" -ForegroundColor Yellow

# Calculate current holiday year (Sept 1 - Aug 31)
$now = Get-Date
$year = $now.Year
if ($now.Month -ge 9) {
    $holidayYear = "$year-$($year + 1)"
} else {
    $holidayYear = "$($year - 1)-$year"
}

Write-Host "Holiday Year: $holidayYear" -ForegroundColor Cyan

# Check if record already exists
Write-Host "`n=== Checking for Existing Record ===" -ForegroundColor Cyan
$filter = "cr_employeeemail eq '$EmployeeEmail' and cr_holidayyear eq '$holidayYear'"
$existingRecords = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($filter))" -Headers $headers

if ($existingRecords.value.Count -gt 0) {
    Write-Host "Record already exists for $EmployeeName in $holidayYear" -ForegroundColor Yellow
    $existing = $existingRecords.value[0]
    Write-Host "  Accrued: $($existing.cr_accrueddays) days"
    Write-Host "  Used: $($existing.cr_useddays) days"
    Write-Host "  Available: $($existing.cr_availabledays) days"

    $update = Read-Host "`nDo you want to update with new dummy data? (yes/no)"
    if ($update -ne "yes") {
        Write-Host "Skipped." -ForegroundColor Yellow
        exit
    }
    $recordId = $existing.cr_holidaybalanceid
    $isUpdate = $true
} else {
    $isUpdate = $false
}

# Create dummy data with realistic values
# Assuming we're mid-year, employee has accrued partial days
$monthsIntoYear = if ($now.Month -ge 9) { $now.Month - 8 } else { $now.Month + 4 }
$accruedDays = [math]::Round($monthsIntoYear * 2.08, 2)  # 2.08 days per month
$usedDays = [math]::Round($accruedDays * 0.4, 1)  # Used about 40% of accrued
$pendingDays = 2  # 2 days pending approval

# Feriefridage (5 per year, accrued monthly)
$feriefridageAccrued = [math]::Round($monthsIntoYear * 0.42, 2)
$feriefridageUsed = [math]::Round($feriefridageAccrued * 0.3, 1)

# Transfer from previous year (simulate 3 days transferred in)
$transferredInDays = 3

# Calculate available
$availableDays = $accruedDays + $transferredInDays - $usedDays - $pendingDays
$feriefridageAvailable = $feriefridageAccrued - $feriefridageUsed

Write-Host "`n=== Dummy Data ===" -ForegroundColor Magenta
Write-Host "  Months into holiday year: $monthsIntoYear"
Write-Host "  Accrued Days: $accruedDays"
Write-Host "  Used Days: $usedDays"
Write-Host "  Pending Days: $pendingDays"
Write-Host "  Transferred In: $transferredInDays"
Write-Host "  Available Days: $availableDays"
Write-Host "  Feriefridage Accrued: $feriefridageAccrued"
Write-Host "  Feriefridage Used: $feriefridageUsed"
Write-Host "  Feriefridage Available: $feriefridageAvailable"

$balanceRecord = @{
    "cr_name" = "$EmployeeName - $holidayYear"
    "cr_employeename" = $EmployeeName
    "cr_employeeemail" = $EmployeeEmail
    "cr_holidayyear" = $holidayYear
    "cr_accrueddays" = $accruedDays
    "cr_useddays" = $usedDays
    "cr_pendingdays" = $pendingDays
    "cr_availabledays" = $availableDays
    "cr_carriedoverdays" = 0
    "cr_transferredindays" = $transferredInDays
    "cr_transferredoutdays" = 0
    "cr_hastransferagreement" = $false
    "cr_feriefridageaccrued" = $feriefridageAccrued
    "cr_feriefridageused" = $feriefridageUsed
    "cr_feriefridageavailable" = $feriefridageAvailable
    "cr_feriefridagetransferredin" = 0
    "cr_feriefridagetransferredout" = 0
    "cr_isactive" = $true
}

Write-Host "`n=== Creating/Updating Record ===" -ForegroundColor Cyan

try {
    if ($isUpdate) {
        # Update existing record
        $updateHeaders = @{
            "Authorization" = "Bearer $token"
            "Content-Type" = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version" = "4.0"
        }
        Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances($recordId)" -Method PATCH -Headers $updateHeaders -Body ($balanceRecord | ConvertTo-Json -Depth 10)
        Write-Host "SUCCESS! Updated holiday balance for $EmployeeName" -ForegroundColor Green
    } else {
        # Create new record
        $response = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances" -Method POST -Headers $headers -Body ($balanceRecord | ConvertTo-Json -Depth 10)
        Write-Host "SUCCESS! Created holiday balance for $EmployeeName" -ForegroundColor Green
        Write-Host "Record ID: $($response.cr_holidaybalanceid)"
    }

    # Verify the record
    Write-Host "`n=== Verifying Record ===" -ForegroundColor Cyan
    $verifyFilter = "cr_employeeemail eq '$EmployeeEmail' and cr_holidayyear eq '$holidayYear'"
    $verifyResponse = Invoke-RestMethod -Uri "$apiUrl/cr_holidaybalances?`$filter=$([System.Uri]::EscapeDataString($verifyFilter))" -Headers $headers

    if ($verifyResponse.value.Count -gt 0) {
        $record = $verifyResponse.value[0]
        Write-Host "`nVerified Holiday Balance:" -ForegroundColor Green
        Write-Host "  Employee: $($record.cr_employeename)"
        Write-Host "  Holiday Year: $($record.cr_holidayyear)"
        Write-Host "  Feriedage:"
        Write-Host "    - Accrued: $($record.cr_accrueddays)"
        Write-Host "    - Used: $($record.cr_useddays)"
        Write-Host "    - Pending: $($record.cr_pendingdays)"
        Write-Host "    - Transferred In: $($record.cr_transferredindays)"
        Write-Host "    - Available: $($record.cr_availabledays)"
        Write-Host "  Feriefridage:"
        Write-Host "    - Accrued: $($record.cr_feriefridageaccrued)"
        Write-Host "    - Used: $($record.cr_feriefridageused)"
        Write-Host "    - Available: $($record.cr_feriefridageavailable)"
    }
} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.Exception.Message
}

Write-Host "`n=== Done ===" -ForegroundColor Green
Write-Host "You can now test the holiday balance card in the SPFx web part!"
