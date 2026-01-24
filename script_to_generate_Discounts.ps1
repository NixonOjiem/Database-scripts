# 1. Install ImportExcel module if you don't have it
# Install-Module ImportExcel -Scope CurrentUser

$excelFile = "D:\shared\DanFolder\diamond\stationary completed.xlsx"
$outputSql = "D:\shared\DanFolder\diamond\stationary completed.sql"
$data = Import-Excel -Path $excelFile

# Start the script by disabling Safe Updates
$sqlStatements = @("SET SQL_SAFE_UPDATES = 0;")

foreach ($row in $data) {
    # Skip rows where barcode might be empty
    if (-not $row.barcode) { continue }

    $barcode = $row.barcode
    $percentage = $row.discount_percentage
    
    # Precise Date Formatting
    $start = [datetime]::Parse($row.discount_start).ToString("yyyy-MM-dd HH:mm:ss")
    $end = [datetime]::Parse($row.discount_end).ToString("yyyy-MM-dd HH:mm:ss")

    $query = "UPDATE products SET discount_percentage = $percentage, discount_start = '$start', discount_end = '$end' WHERE barcode = '$barcode';"
    $sqlStatements += $query
}

# Re-enable Safe Updates at the end
$sqlStatements += "SET SQL_SAFE_UPDATES = 1;"

# Write to a file
$sqlStatements | Out-File -FilePath $outputSql -Encoding ascii
Write-Host "Success! SQL script generated: $outputSql" -ForegroundColor Green