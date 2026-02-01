# Analyze the Ferieregistrering CSV data

$csvPath = "c:\Users\info\Innovater\absence-registration\scripts\Ferieregistrering.csv"
$data = Import-Csv $csvPath -Encoding UTF8

Write-Host "=== Data Summary ===" -ForegroundColor Magenta
Write-Host "Total records: $($data.Count)"
Write-Host "Unique employees: $(($data | Select-Object -ExpandProperty Medarbejder -Unique).Count)"

Write-Host "`n=== Categories ===" -ForegroundColor Magenta
$data | Group-Object Kategori | Select-Object Name, Count | Format-Table -AutoSize

Write-Host "`n=== Date Range ===" -ForegroundColor Magenta
$dates = $data | ForEach-Object {
    try {
        [DateTime]::ParseExact($_.'Dato start', 'dd-MM-yyyy', $null)
    } catch {
        $null
    }
} | Where-Object { $_ -ne $null }
$dateStats = $dates | Measure-Object -Minimum -Maximum
Write-Host "From: $($dateStats.Minimum.ToString('yyyy-MM-dd'))"
Write-Host "To: $($dateStats.Maximum.ToString('yyyy-MM-dd'))"

Write-Host "`n=== Sample Employees ===" -ForegroundColor Magenta
$data | Select-Object -ExpandProperty Medarbejder -Unique | Select-Object -First 10
