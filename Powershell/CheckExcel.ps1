
#Generated with the help of grok
#Performance tuning done outside of Grok
clear
# Define paths
$excelFilePath = "C:\Temp\Specs.xlsx"  # Change this to your Excel file path
$outputFilePath = "C:\Temp\StrikethroughResults.txt"  # Change this to your output file path
$logpath  = "C:\Temp\ExcelChklog.log"  # Change this to your output file path

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open the workbook
    $workbook = $excel.Workbooks.Open($excelFilePath)
    
    # Initialize output array
    $results = @()
    $log = @()

    Write-Host "Start checking the Excel file"
    
    # Loop through each worksheet
    foreach ($sheet in $workbook.Worksheets) {
        $sheetName = $sheet.Name
        $usedRange = $sheet.UsedRange
        
        $rowcnt = $usedRange.Rows.Count
        $colcnt = $usedRange.Columns.Count
        $log += "Sheet: " + $sheetName.ToString().PadRight(10) + " #Rows: " + $rowcnt.ToString().PadRight(5) + " #Columns: " + $colcnt.ToString().PadRight(5)
        Write-Host "Sheet: "$sheetName.ToString().PadRight(10) " #Rows: "$rowcnt.ToString().PadRight(5) " #Columns: "$colcnt.ToString().PadRight(5)
        
        # Loop through each row in the used range
        for ($row = 1; $row -le $rowcnt; $row++) {
            #If there's no data on row 1 skip the check
            if ($sheet.Cells.Item($row, 1).Text) {
                # Loop through each column in the row
                for ($col = 1; $col -le $colcnt; $col++) {
                    $cell = $sheet.Cells.Item($row, $col)
                    # Check if cell has content and for performance then check for strikethrough formatting
                    if ($sheet.Cells.Item($row, $col).Text)  { 
                    $cell = $sheet.Cells.Item($row, $col)
                      if ($cell.Font.Strikethrough) {
                        $results += "Sheet: $sheetName, Row: $row, Column: $col - Strikethrough found"
                      }
                    }
                }
            }
        }
    }
    
    # If results found, write to file
    if ($results.Count -gt 0) {
        $results | Out-File -FilePath $outputFilePath -Encoding UTF8
        Write-Host "Strikethrough text found. Results written to: $outputFilePath"
    } else {
        "No strikethrough text found in the Excel file." | Out-File -FilePath $outputFilePath -Encoding UTF8
        Write-Host "No strikethrough text found."
    }
    
    $log | Out-File -FilePath $logpath -Encoding UTF8

}
catch {
    # Error handling
    "Error occurred: $_" | Out-File -FilePath $outputFilePath -Encoding UTF8
    Write-Host "An error occurred: $_"
}
finally {
    # Improved cleanup with error handling and delays
    try {
        if ($workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
    }
    catch {
        Write-Host "Warning: Failed to close workbook - $_"
    }

    try {
        if ($excel) {
            $excel.Quit()
            Start-Sleep -Milliseconds 500  # Give Excel time to shut down
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
    catch {
        Write-Host "Warning: Failed to quit Excel - $_"
    }

    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Optional: Kill Excel process if still running (uncomment if needed)
    # Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
}