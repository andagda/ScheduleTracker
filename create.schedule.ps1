
# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# Add a new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Define the headers
$daysOfWeek = @("Su", "M", "T", "W", "Th", "F", "Sa")
$values = "APE,OS,PTO,PTH,OB,H"
$AH=34
$AI=35
$AJ=36
$AK=37
$AL=38
$AM=39
$AN=40

# Get the current directory
$currentDirectory = Get-Location
$currentYear = (Get-Date).Year
$filePath= "$currentDirectory\ScheduleTracker_$currentYear.xlsx"

if (Test-Path $filePath) {
    Remove-Item -Path $filePath
}

# Loop through each month of the year
for ($month = 1; $month -le 12; $month++) {
    
    $daysInMonth = [DateTime]::DaysInMonth($currentYear, $month)
    $monthName = (Get-Date -Year $currentYear -Month $month -Day 1).ToString("MMMM")
    
    # Calculate the starting row for each month's table
    $startRow = ($month - 1) * 10 + 1

    # Merge cells for the month name header
    $worksheet.Cells.Item($startRow, 1).Value = $monthName
    $worksheet.Cells.Item($startRow, 1).HorizontalAlignment = -4108  # Center alignment
    $worksheet.Cells.Item($startRow, 1).Font.Bold = $true
    $worksheet.Cells.Item($startRow, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Blue)  # Set background color
    $worksheet.Cells.Item($startRow, 1).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
    $worksheet.Cells.Item($startRow+1, 1).Value = "Name"
    $worksheet.Cells.Item($startRow+1, 1).HorizontalAlignment = -4108  # Center alignment
    $worksheet.Cells.Item($startRow+1, 1).Font.Bold = $true
    $worksheet.Cells.Item($startRow+1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Green)  # Set background color
    $worksheet.Cells.Item($startRow+1, 1).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
    
    # Set the main headers (dates of the month)
    for ($i = 1; $i -le $daysInMonth; $i++) {
        $date = Get-Date -Year $currentYear -Month $month -Day $i
        $worksheet.Cells.Item($startRow, $i + 1) = $date.ToString("dd")
        $worksheet.Cells.Item($startRow + 1, $i + 1) = $daysOfWeek[$date.DayOfWeek.value__]
        if ($daysOfWeek[$date.DayOfWeek.value__] -eq "Sa" -or $daysOfWeek[$date.DayOfWeek.value__] -eq "Su") {
            $worksheet.Cells.Item($startRow +1 , $i + 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
        }
        # I want to set the column width to 5 pixels for each day
        $worksheet.Columns.Item($i + 1).ColumnWidth = 5
    }
    
    # Set the secondary header always 
    $worksheet.Cells.Item($startRow, $AH) = "Working Days"
    $worksheet.Cells.Item($startRow, $AH).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $AH).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)  # Set background color
    $worksheet.Cells.Item($startRow, $AI) = "Holidays"
    $worksheet.Cells.Item($startRow, $AI).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $AI).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)  # Set font color

    $worksheet.Cells.Item($startRow, $AJ) = "OS"
    $lastColumn = $worksheet.Cells.Item($startRow + 1, $daysInMonth + 1).Address(0, 0)
    $worksheet.Cells.Item($startRow + 2, $AJ).Formula = "=COUNTIF(B$($startRow + 3):$lastColumn, `"OS`")"
    $worksheet.Cells.Item($startRow, $AK) = "PTO"
    $worksheet.Cells.Item($startRow + 2, $AK).Formula = "=COUNTIF(B$($startRow + 3):$lastColumn, `"PTO`")"
    $worksheet.Cells.Item($startRow, $AL) = "PTH"
    $worksheet.Cells.Item($startRow + 2, $AL).Formula = "=COUNTIF(B$($startRow + 3):$lastColumn, `"PTH`")/2"
    $worksheet.Cells.Item($startRow, $AM) = "OB"
    $worksheet.Cells.Item($startRow + 2, $AM).Formula = "=COUNTIF(B$($startRow + 3):$lastColumn, `"OB`")"
    $worksheet.Cells.Item($startRow, $AN) = "APE"
    $worksheet.Cells.Item($startRow + 2, $AN).Formula = "=COUNTIF(B$($startRow + 3):$lastColumn, `"APE`")/2"

    
    # Create the drop-down list for the main data column
    for ($i = $startRow + 2; $i -le $startRow + 9; $i++) {
        for ($j = 2; $j -le $daysInMonth + 1; $j++) {
            $cell = $worksheet.Cells.Item($i, $j)
            $validation = $cell.Validation
            $validation.Delete()
            $validation.Add(3, 1, 1, $values)
            $validation.IgnoreBlank = $true
            $validation.InCellDropdown = $true
        }
    }
}


# Add conditional formatting
$range = $worksheet.Range("B$($startRow + 2):$lastColumn$($startRow + 9)")
$formatConditions = $range.FormatConditions
$conditionPTH = $formatConditions.Add(1, 3, "=`"PTH`"") # 1 = xlCellValue, 3 = xlEqual
$conditionPTH.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Turquoise)
$conditionPTO = $formatConditions.Add(1, 3, "=`"PTO`"") # 1 = xlCellValue, 3 = xlEqual
$conditionPTO.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Orange)


# Save the workbook

$workbook.SaveAs($filePath)
$workbook.Close()
$excel.Quit()

# Release the COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null