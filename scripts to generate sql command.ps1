# Import data from Excel and generate SQL
$excelPath = "D:\oita.xlsx"
$outputSqlFile = "D:\oita.sql"

# Install ImportExcel module if needed
if (-not (Get-Module -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

$data = Import-Excel -Path $excelPath

$valueClauses = foreach ($row in $data) {
    # Generate UUID in PowerShell
    $uuid = [guid]::NewGuid().ToString()
    
    # Handle empty values properly
    $barcode = if ($row.barcode) { "'$($row.barcode)'" } else { "NULL" }
    $title = if ($row.title) { "'$($row.title.Replace("'","''"))'" } else { "NULL" }
    $description = if ($row.description) { "'$($row.description.Replace("'","''"))'" } else { "NULL" }
    $selling_price = if ($row.selling_price) { $row.selling_price.ToString() } else { "NULL" }
    $category_id = if ($row.category_id) { $row.category_id.ToString() } else { "NULL" }
    $sub_category = if ($row.sub_category) { $row.sub_category.ToString() } else { "NULL" }
    $mini_category = if ($row.mini_category) { $row.mini_category.ToString() } else { "NULL" }
    $brand_id = if ($row.brand_id) { $row.brand_id.ToString() } else { "NULL" }
	$product_type = if ($row.product_type) { $row.product_type.ToString() } else { "NULL" }
    $product_image = if ($row.product_image) { "'$($row.product_image).jpg'" } else { "NULL" }
    $product_image_1 = if ($row.product_image_1) { "'$($row.product_image_1).jpg'" } else { "NULL" }
    $product_image_2 = if ($row.product_image_2) { "'$($row.product_image_2).jpg'" } else { "NULL" }
	$product_image_3 = if ($row.product_image_3) { "'$($row.product_image_3).jpg'" } else { "NULL" }
	$product_image_4 = if ($row.product_image_4) { "'$($row.product_image_4).jpg'" } else { "NULL" }
	$product_image_5 = if ($row.product_image_5) { "'$($row.product_image_5).jpg'" } else { "NULL" }
	$product_image_6 = if ($row.product_image_6) { "'$($row.product_image_6).jpg'" } else { "NULL" }
	$product_image_7 = if ($row.product_image_7) { "'$($row.product_image_7).jpg'" } else { "NULL" }
	$product_image_8 = if ($row.product_image_8) { "'$($row.product_image_8).jpg'" } else { "NULL" }
	$product_image_9 = if ($row.product_image_9) { "'$($row.product_image_9).jpg'" } else { "NULL" }
	$product_image_10 = if ($row.product_image_10) { "'$($row.product_image_10).jpg'" } else { "NULL" }
    $quantity = if ($row.quantity) { $row.quantity.ToString() } else { "NULL" }
    $active = if ($row.active) { $row.active.ToString() } else { "NULL" }

@"
        (
            '$uuid',
            $barcode,
            $title,
            $description,
            $selling_price,
            $category_id,
            $sub_category,
            $mini_category,
            $brand_id,
			$product_type,
            $product_image,
            $product_image_1,
            $product_image_2,
			$product_image_3,
			$product_image_4,
			$product_image_5,
			$product_image_6,
			$product_image_7,
			$product_image_8,
			$product_image_9,
			$product_image_10,
            $quantity,
            $active,
            NOW(),
            NOW()
        )
"@
}

$sqlCommand = @"
-- SET SESSION innodb_autoinc_lock_mode = 0;
ALTER TABLE products AUTO_INCREMENT = 2953;
INSERT IGNORE INTO products 
    (uuid, barcode, title, description, selling_price,
    category_id, sub_category, mini_category, brand_id, product_type,
    product_image, product_image_1, product_image_2, product_image_3, product_image_4,product_image_5,
	product_image_6,product_image_7,product_image_8,product_image_9,product_image_10,
    quantity, active, created_at, updated_at)
VALUES
$($valueClauses -join ",`n");


-- SET SESSION innodb_autoinc_lock_mode = 1;
"@

$sqlCommand | Out-File -FilePath $outputSqlFile -Encoding utf8
Write-Host "SQL file generated: $outputSqlFile"