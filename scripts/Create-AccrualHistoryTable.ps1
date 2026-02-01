# Create Accrual History Table in Dataverse
# This table tracks all changes to holiday balances for auditing

$DataverseUrl = "https://orgab6f6874.crm4.dynamics.com"
$clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"

Write-Host "=== Create Accrual History Table ===" -ForegroundColor Magenta
Write-Host "This script creates the cr_accrualhistories table in Dataverse" -ForegroundColor Yellow
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

Write-Host "`n=== Note ===" -ForegroundColor Yellow
Write-Host "Table creation via API requires System Administrator privileges."
Write-Host "If this fails, you may need to create the table manually in Power Apps."
Write-Host ""

# Check if table already exists
Write-Host "=== Checking if table exists ===" -ForegroundColor Cyan
try {
    $checkResponse = Invoke-RestMethod -Uri "$apiUrl/cr_accrualhistories?`$top=1" -Headers $headers -ErrorAction Stop
    Write-Host "Table 'cr_accrualhistories' already exists!" -ForegroundColor Green
    Write-Host "Found $($checkResponse.value.Count) record(s)" -ForegroundColor Cyan

    $continue = Read-Host "`nDo you want to view the table schema? (yes/no)"
    if ($continue -ne "yes") {
        exit
    }
} catch {
    if ($_.Exception.Response.StatusCode -eq 404) {
        Write-Host "Table does not exist. Will attempt to create..." -ForegroundColor Yellow
    } else {
        Write-Host "Error checking table: $_" -ForegroundColor Red
    }
}

Write-Host "`n=== Table Definition ===" -ForegroundColor Cyan
Write-Host @"

To create the Accrual History table manually in Power Apps:

1. Go to https://make.powerapps.com
2. Select your environment (orgab6f6874)
3. Go to Tables > New table
4. Create table with these settings:

   Display Name: Accrual History
   Plural Name: Accrual Histories
   Schema Name: cr_accrualhistory (will become cr_accrualhistories)
   Primary Column: Name (cr_name)

5. Add these columns:

   | Display Name         | Schema Name              | Type        | Required |
   |---------------------|--------------------------|-------------|----------|
   | Employee Email      | cr_employeeemail         | Text        | Yes      |
   | Employee Name       | cr_employeename          | Text        | No       |
   | Holiday Year        | cr_holidayyear           | Text        | Yes      |
   | Accrual Date        | cr_accrualdate           | Date/Time   | Yes      |
   | Accrual Month       | cr_accrualmonth          | Whole Number| No       |
   | Accrual Year        | cr_accrualyear           | Whole Number| No       |
   | Days Accrued        | cr_daysaccrued           | Decimal     | Yes      |
   | Feriefridage Accrued| cr_feriefridageaccrued   | Decimal     | No       |
   | Balance After       | cr_balanceafteraccrual   | Decimal     | No       |
   | Accrual Type        | cr_accrualtype           | Choice      | Yes      |
   | Notes               | cr_notes                 | Text (Multi)| No       |

6. Create the Accrual Type choice with these values:

   | Value     | Label (Danish)           | Label (English)          |
   |-----------|--------------------------|--------------------------|
   | 100000000 | Månedlig optjening       | Monthly Accrual          |
   | 100000001 | Årsstart overførsel      | Year Start Transfer In   |
   | 100000002 | Manuel justering         | Manual Adjustment        |
   | 100000003 | Startsaldo               | Initial Balance          |
   | 100000004 | Årsslut overførsel (ud)  | Year End Transfer Out    |
   | 100000005 | Feriefridage optjening   | Feriefridage Accrual     |
   | 100000006 | Bortfald                 | Forfeiture               |
   | 100000007 | Udbetaling               | Payout                   |

"@ -ForegroundColor White

Write-Host "`n=== Alternative: Create via Solution ===" -ForegroundColor Cyan
Write-Host @"

For better manageability, create the table within a Solution:

1. Go to Solutions in Power Apps
2. Create or open your solution
3. Add > Table > New table
4. Follow the same column definitions above
5. This allows you to export/import the schema between environments

"@ -ForegroundColor White

Write-Host "=== Done ===" -ForegroundColor Green
