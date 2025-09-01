param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputSqlPath
)

# Import the module
Import-Module ImportExcel

# Read Excel data
$data = Import-Excel -Path $ExcelPath

# Create SQL file
$null = New-Item -Path $OutputSqlPath -Force -ItemType File

foreach ($row in $data) {
    $columns = @()
    $values = @()

    # Process each column
    $row.PSObject.Properties | ForEach-Object {
        $colName = $_.Name
        $val = $_.Value
        
        if ($null -ne $colName) {
            $columns += $colName
            
            if ($null -eq $val -or $val -eq "") {
                $values += "NULL"
            }
            else {
                switch ($colName) {
                    { $_ -in "id", "barcode", "selling_price", "category_id", "sub_category",
                            "mini_category", "brand_id", "product_type", "quantity", "active" } {
                        $values += $val
                        break
                    }
                    
                    { $_ -in "created_at", "updated_at" } {
                        if ($val -is [DateTime]) {
                            $values += "'$($val.ToString("yyyy-MM-dd HH:mm:ss"))'"
                        }
                        else {
                            $values += "'$val'"
                        }
                        break
                    }
                    
                    default {
                        $escaped = $val -replace "'", "''"
                        $values += "'$escaped'"
                    }
                }
            }
        }
    }

    $columnList = $columns -join ", "
    $valueList = $values -join ", "

    "INSERT INTO products ($columnList) VALUES ($valueList);" | Add-Content -Path $OutputSqlPath
}

Write-Host "SQL file generated at $OutputSqlPath with $($data.Count) records"