# Configuration
$folderPath = "D:\shared\Barcodes\Clothes"
$excelPath =  "D:\shared\Barcodes\Clothes\clothes.xlsx"

# 1. Verify Excel installation
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
}
catch {
    Write-Host "Error: Microsoft Excel not installed!" -ForegroundColor Red
    exit
}

# 2. Validate folder and files
if (-not (Test-Path $folderPath)) {
    Write-Host "Folder not found: $folderPath" -ForegroundColor Red
    exit
}

# Modified file detection with case-insensitive check
$images = Get-ChildItem -Path $folderPath -File |
          Where-Object { $_.Extension -match '\.(jpg|jpeg|png|gif|bmp)$' -and $_.Attributes -notmatch 'Hidden|System' }

# Debug output
Write-Host "Detected files:" -ForegroundColor Cyan
$images | Format-Table Name, Extension, Length

if (-not $images) {
    Write-Host "No images found in folder!" -ForegroundColor Yellow
    Write-Host "Check these aspects:" -ForegroundColor Cyan
    Write-Host "1. File extensions (actual extensions: $((Get-ChildItem $folderPath).Extension | Select-Object -Unique))"
    Write-Host "2. Hidden/system files"
    Write-Host "3. File permissions"
    $excel.Quit()
    exit
}

# Rest of the script remains the same...

# 3. Excel processing with proper cleanup
try {
    $excel.Visible = $false  # Set to $true to watch the process
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)

    # Create headers
    $worksheet.Cells.Item(1, 1) = "Image Name"
    $worksheet.Cells.Item(1, 2) = "Extension"
    $worksheet.Cells.Item(1, 3) = "File Size (KB)"
    $worksheet.Cells.Item(1, 4) = "Created Date"

    # Format headers
    $headerRange = $worksheet.Range("A1:D1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 15773696  # Light blue header
    $headerRange.HorizontalAlignment = -4108  # Center align

    # Populate data
    $row = 2
    foreach ($img in $images) {
        $worksheet.Cells.Item($row, 1) = $img.BaseName
        $worksheet.Cells.Item($row, 2) = $img.Extension.TrimStart('.')
        $worksheet.Cells.Item($row, 3) = [math]::Round($img.Length/1KB, 2)
        $worksheet.Cells.Item($row, 4) = $img.CreationTime.ToString("yyyy-MM-dd")
        $row++
    }

    # Auto-fit columns and apply borders
    $usedRange = $worksheet.UsedRange
    $usedRange.Columns.AutoFit() | Out-Null
    $usedRange.Borders.LineStyle = 1  # Continuous borders

    # Save and close
    $workbook.SaveAs($excelPath)
    Write-Host "Successfully created: $excelPath" -ForegroundColor Green
}
catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
}
finally {
    # Mandatory cleanup
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Force kill any remaining Excel processes
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
}