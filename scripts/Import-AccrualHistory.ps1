# Import-AccrualHistory.ps1
# Imports only accrual records (Tilskrivning) into cr_accrualhistory

param(
    [string]$CsvPath = "c:\Users\info\Innovater\absence-registration\scripts\Ferieregistrering.csv",
    [string]$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com",
    [int]$StartYear = 2025
)

$HOLIDAY_YEAR_START_MONTH = 9

function Get-HolidayYear {
    param([DateTime]$Date)
    if ($Date.Month -ge $HOLIDAY_YEAR_START_MONTH) {
        return "$($Date.Year)-$($Date.Year + 1)"
    } else {
        return "$($Date.Year - 1)-$($Date.Year)"
    }
}

# Authenticate
Write-Host "=== Authenticating ===" -ForegroundColor Cyan
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"
$scope = "$DataverseUrl/.default"

$deviceCodeResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode" -Method POST -Body @{ client_id = $clientId; scope = $scope }
Write-Host $deviceCodeResponse.message -ForegroundColor Yellow

$token = $null
$startTime = Get-Date
while ((Get-Date) -lt $startTime.AddSeconds($deviceCodeResponse.expires_in)) {
    Start-Sleep -Seconds $deviceCodeResponse.interval
    try {
        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/organizations/oauth2/v2.0/token" -Method POST -Body @{
            grant_type = "urn:ietf:params:oauth:grant-type:device_code"
            client_id = $clientId
            device_code = $deviceCodeResponse.device_code
        }
        $token = $tokenResponse.access_token
        Write-Host "Authenticated!" -ForegroundColor Green
        break
    } catch { }
}

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}
$apiUrl = "$DataverseUrl/api/data/v9.2"

# Load and filter data
Write-Host "`n=== Loading Data ===" -ForegroundColor Cyan
$allData = Import-Csv $CsvPath -Encoding UTF8

# Filter: Only accrual records from StartYear onwards
$data = $allData | Where-Object {
    $cat = $_.Kategori
    if ($cat -ne "Tilskrivning, feriedage" -and $cat -ne "Tilskrivning, feriefridage") { return $false }
    try {
        $date = [DateTime]::ParseExact($_.'Dato start', 'dd-MM-yyyy', $null)
        return $date.Year -ge $StartYear
    } catch { return $false }
}

Write-Host "Found $($data.Count) accrual records from $StartYear onwards"

# Import
Write-Host "`n=== Importing ===" -ForegroundColor Cyan
$success = 0; $errors = 0

foreach ($row in $data) {
    $date = [DateTime]::ParseExact($row.'Dato start', 'dd-MM-yyyy', $null)
    $days = [decimal]($row.'Antal dage' -replace ',', '.')
    $isFeriefridage = $row.Kategori -eq "Tilskrivning, feriefridage"

    $record = @{
        "cr_name" = "$($row.Medarbejder) - $(Get-Date $date -Format 'MMM yyyy')"
        "cr_employeename" = $row.Medarbejder
        "cr_employeeemail" = ""
        "cr_holidayyear" = Get-HolidayYear -Date $date
        "cr_accrualdate" = $date.ToString("yyyy-MM-dd")
        "cr_accrualmonth" = $date.Month
        "cr_accrualyear" = $date.Year
        "cr_daysaccrued" = if ($isFeriefridage) { 0 } else { $days }
        "cr_feriefridageaccrued" = if ($isFeriefridage) { $days } else { $null }
        "cr_accrualtype" = 100000000  # Monthly Accrual
        "cr_notes" = "Import: $($row.Kategori)"
    }

    try {
        Invoke-RestMethod -Uri "$apiUrl/cr_accrualhistories" -Method POST -Headers $headers -Body ($record | ConvertTo-Json) | Out-Null
        $success++
        if ($success % 25 -eq 0) { Write-Host "  $success imported..." -ForegroundColor Gray }
    } catch {
        $errors++
        if ($errors -le 3) { Write-Warning "Error: $($row.Medarbejder) - $_" }
    }
}

Write-Host "`n=== Done ===" -ForegroundColor Green
Write-Host "Success: $success | Errors: $errors"
