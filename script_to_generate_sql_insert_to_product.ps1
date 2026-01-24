# Import data from Excel and generate SQL
$excelPath = "D:\shared\Barcodes\Stationery\complete.xlsx"
$outputSqlFile = "D:\shared\Barcodes\Stationery\complete.sql"

# Install ImportExcel module if needed
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

$data = Import-Excel -Path $excelPath

$valueClauses = foreach ($row in $data) {
    # Generate UUID
    $uuid = [guid]::NewGuid().ToString()
    
    # Handle existing values and escaping strings
    $barcode = if ($row.barcode) { "'$($row.barcode)'" } else { "NULL" }
    $title = if ($row.title) { "'$($row.title.Replace("'","''"))'" } else { "NULL" }
    $description = if ($row.description) { "'$($row.description.Replace("'","''"))'" } else { "NULL" }
    $selling_price = if ($row.selling_price) { $row.selling_price.ToString() } else { "NULL" }
    $category_id = if ($row.category_id) { $row.category_id.ToString() } else { "NULL" }
    $sub_category = if ($row.sub_category) { $row.sub_category.ToString() } else { "NULL" }
    $mini_category = if ($row.mini_category) { $row.mini_category.ToString() } else { "NULL" }
    $brand_id = if ($row.brand_id) { $row.brand_id.ToString() } else { "NULL" }
    $product_type = if ($row.product_type) { $row.product_type.ToString() } else { "NULL" }
    
    # Image handling
    $product_image = if ($row.product_image) { "'$($row.product_image).jpg'" } else { "NULL" }
    $img1 = if ($row.product_image_1) { "'$($row.product_image_1).jpg'" } else { "NULL" }
    $img2 = if ($row.product_image_2) { "'$($row.product_image_2).jpg'" } else { "NULL" }
    $img3 = if ($row.product_image_3) { "'$($row.product_image_3).jpg'" } else { "NULL" }
    $img4 = if ($row.product_image_4) { "'$($row.product_image_4).jpg'" } else { "NULL" }
    $img5 = if ($row.product_image_5) { "'$($row.product_image_5).jpg'" } else { "NULL" }
    $img6 = if ($row.product_image_6) { "'$($row.product_image_6).jpg'" } else { "NULL" }
    $img7 = if ($row.product_image_7) { "'$($row.product_image_7).jpg'" } else { "NULL" }
    $img8 = if ($row.product_image_8) { "'$($row.product_image_8).jpg'" } else { "NULL" }
    $img9 = if ($row.product_image_9) { "'$($row.product_image_9).jpg'" } else { "NULL" }
    $img10 = if ($row.product_image_10) { "'$($row.product_image_10).jpg'" } else { "NULL" }
    
    $quantity = if ($row.quantity) { $row.quantity.ToString() } else { "NULL" }
    $active = if ($row.active) { $row.active.ToString() } else { "NULL" }
    
    # New Fields and Meta Keywords
    $meta_keywords = if ($row.meta_keywords) { "'$($row.meta_keywords.Replace("'","''"))'" } else { "NULL" }
    $discount_percentage = "0"    # Default 0
    $discount_start = "NULL"      # Default NULL
    $discount_end = "NULL"        # Default NULL

@"
        (
            '$uuid', $barcode, $title, $description, $selling_price,
            $category_id, $sub_category, $mini_category, $brand_id, $product_type,
            $product_image, $img1, $img2, $img3, $img4, $img5, $img6, $img7, $img8, $img9, $img10,
            $quantity, $active, $meta_keywords, NOW(), NOW(),
            $discount_percentage, $discount_start, $discount_end
        )
"@
}

$sqlCommand = @"
-- Set the increment and insert data
ALTER TABLE products AUTO_INCREMENT = 3951;
INSERT IGNORE INTO products 
    (uuid, barcode, title, description, selling_price,
    category_id, sub_category, mini_category, brand_id, product_type,
    product_image, product_image_1, product_image_2, product_image_3, product_image_4, product_image_5,
    product_image_6, product_image_7, product_image_8, product_image_9, product_image_10,
    quantity, active, meta_keywords, created_at, updated_at,
    discount_percentage, discount_start, discount_end)
VALUES
$($valueClauses -join ",`n");
"@

# FIX: Use [System.IO.File] to write the file without a Byte Order Mark (BOM)
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($outputSqlFile, $sqlCommand, $utf8NoBom)

Write-Host "SQL file generated successfully without BOM: $outputSqlFile"