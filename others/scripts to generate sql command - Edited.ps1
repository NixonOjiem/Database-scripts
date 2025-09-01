# Importing CSV data and generating MySQL INSERT statement
$csvFile = "D:\products.csv"  # Update with actual CSV file path
$outputFile = "D:\Products-G.sql"

# Read CSV data
$rows = Import-Csv -Path $csvFile

# Start the bulk INSERT statement
$sql = "INSERT INTO products ( uuid, barcode, title, description, selling_price, category_id, sub_category, mini_category, brand_id, product_image, product_image_1, product_image_2, product_image_3, product_image_4, product_image_5, product_image_6, product_image_7, product_image_8, product_image_9, product_image_10, quantity, active, meta_keywords, created_at, updated_at) VALUES"

# Iterate through each row and format values
$values = @()

foreach ($row in $rows) {
    $created_at = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $updated_at = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Escape single quotes for SQL syntax
    $title = $row.title -replace "'", "''"
    $description = $row.description -replace "'", "''"
    $meta_keywords = $row.meta_keywords -replace "'", "''"

    # Format values for SQL
    $values += "('$($row.uuid)', '$($row.barcode)', '$title', '$description', '$($row.selling_price)', '$($row.category_id)', '$($row.sub_category)', '$($row.mini_category)', '$($row.brand_id)', '$($row.product_image)', '$($row.product_image_1)', '$($row.product_image_2)','$($row.product_image_3)','$($row.product_image_4)', '$($row.product_image_5)','$($row.product_image_6)','$($row.product_image_7)','$($row.product_image_8)','$($row.product_image_9)','$($row.product_image_10)', '$($row.quantity)', '$($row.active)', '$meta_keywords', '$created_at', '$updated_at')"
}

# Combine values into a single SQL statement
$sql += " " + ($values -join ", ") + ";"

# Write to file
$sql | Out-File -Encoding UTF8 -FilePath $outputFile

Write-Host "Bulk SQL insert statement saved to $outputFile"
