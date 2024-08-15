
<#
.SYNOPSIS
Series of commands that helps build the CheckinApp to local server

.DESCRIPTION
  Performs building and running the CheckinApp. 

.PARAMETER year
  -$year -branchName parameter will pull and checkout desired branch.

.EXAMPLE 
    .\create.schedule.ps1 -year 2025
    ./create.schedule.ps1  

.NOTES
/*==================================================================================================
 = This file is part of the Navitaire CheckinApp application.
 = Copyright © Navitaire LLC, an Amadeus company. All rights reserved.
 =================================================================================================*/
#>

<# 
.PARAMETERS EXPLANATION:
year: Year that you want the schdule tracker to be created. This will create table from January to December When blank it will default to the current year.
#>

param (
    [int]$year = (Get-Date).Year,
    [int]$teamsize = 8
)

Write-Host "`n`nCreating ScheduleTracker_$year.xlsx for a team of $teamsize.........." -ForegroundColor Blue

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# Add a new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Define the headers
$daysOfWeek = @("Su", "M", "T", "W", "Th", "F", "Sa")
$values = "APE,OS,PTO,PTH,OB,H"
$columnMapping = @{
    30 = "AD"
    31 = "AE"
    32 = "AF"
    33 = "AG"
    34 = "AH"
    35 = "AI"
    36 = "AJ"
    37 = "AK"
    38 = "AL"
    39 = "AM"
    40 = "AN"
    41 = "AO"
}

# Create a reverse mapping hashtable to map string values to integer representations
$reverseColumnMapping = @{}
foreach ($key in $columnMapping.Keys) {
    $reverseColumnMapping[$columnMapping[$key]] = $key
}
$weekdayColumnValue=$reverseColumnMapping["AG"]
$holidayColumnValue=$reverseColumnMapping["AH"]
$workingDaysColumnValue=$reverseColumnMapping["AI"]
$OSColumnValue = $reverseColumnMapping["AJ"]
$PTOColumnValue = $reverseColumnMapping["AK"]
$PTHColumnValue = $reverseColumnMapping["AL"]
$OBColumnValue = $reverseColumnMapping["AM"]
$APEColumnValue = $reverseColumnMapping["AN"]
$percentColumnValue = $reverseColumnMapping["AO"]

function SetFormulaHeaders ($startRow, $lastCol) {
    $nextRow = $startRow + 1
    # Set the headers for columns with formulas 
    $worksheet.Cells.Item($startRow, $weekdayColumnValue) = "Weekdays"
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
    # Set the formula for weekdays
    $worksheet.Cells.Item($nextRow, $weekdayColumnValue).Formula = "=COUNTIF(B$($startRow + 1):AF$($startRow+1), `"M`") + COUNTIF(B$($startRow + 1):AF$($startRow+1), `"T`") + COUNTIF(B$($startRow + 1):AF$($startRow+1), `"W`") + COUNTIF(B$($startRow + 1):AF$($startRow+1), `"Th`") + COUNTIF(B$($startRow + 1):AF$($startRow+1), `"F`")"
    $worksheet.Cells.Item($startRow, $holidayColumnValue) = "Holidays"
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)  # Set font color
    # Set the formula for holidays
    $worksheet.Cells.Item($nextRow, $holidayColumnValue).Formula = "=COUNTIF(B$($startRow + 3):AF$($startRow+3), `"H`")"
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue) = "Working Days"
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)  # Set background color
    # Set the formula for working days
    $worksheet.Cells.Item($nextRow, $workingDaysColumnValue).Formula = "=AG$nextRow - AH$nextRow"
    $worksheet.Cells.Item($nextRow, $OSColumnValue) = "OS"
    $worksheet.Cells.Item($nextRow, $PTOColumnValue) = "PTO"
    $worksheet.Cells.Item($nextRow, $PTHColumnValue) = "PTH"
    $worksheet.Cells.Item($nextRow, $OBColumnValue) = "OB"
    $worksheet.Cells.Item($nextRow, $APEColumnValue) = "APE"
    $worksheet.Cells.Item($nextRow, $percentColumnValue) = "%"
    $worksheet.Cells.Item($startRow, $percentColumnValue).HorizontalAlignment = -4108  # Center alignment 
    $worksheet.Cells.Item($startRow, $percentColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $percentColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkGreen)  # Set font color
}
function SetExcelFormulas ($startRow, $lastCol) {
    $worksheet.Cells.Item($startRow, $OSColumnValue).Formula = "=COUNTIF(B$($startRow + 3):$lastCol, `"OS`")"
    $worksheet.Cells.Item($startRow, $PTOColumnValue).Formula = "=COUNTIF(B$($startRow + 3):$lastCol, `"PTO`")"
    $worksheet.Cells.Item($startRow, $PTHColumnValue).Formula = "=COUNTIF(B$($startRow + 3):$lastCol, `"PTH`")/2"
    $worksheet.Cells.Item($startRow, $OBColumnValue).Formula = "=COUNTIF(B$($startRow + 3):$lastCol, `"OB`")"
    $worksheet.Cells.Item($startRow, $APEColumnValue).Formula = "=COUNTIF(B$($startRow + 3):$lastCol, `"APE`")/2"
}

# Get the current directory
$currentDirectory = Get-Location
$filePath= "$currentDirectory\ScheduleTracker_$year.xlsx"

if (Test-Path $filePath) {
    Remove-Item -Path $filePath
}

# Loop through each month of the year that was specified
for ($month = 1; $month -le 12; $month++) {
    
    
    $daysInMonth = [DateTime]::DaysInMonth($year, $month)
    $monthName = (Get-Date -Year $year -Month $month -Day 1).ToString("MMMM")
    Write-Host "Generating Table for $monthName" -ForegroundColor Cyan
    
    # Calculate the starting row for each month's table
    $startRow = ($month - 1) * ($teamsize + 2) + 1

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
    
    # hash table to store the weekend columns
    $weekendColumns = @()

    # Set the main headers (dates of the month)
    for ($i = 1; $i -le $daysInMonth; $i++) {
        $date = Get-Date -Year $year -Month $month -Day $i
        $worksheet.Cells.Item($startRow, $i + 1) = $date.ToString("dd")
        $worksheet.Cells.Item($startRow + 1, $i + 1) = $daysOfWeek[$date.DayOfWeek.value__]
        if ($daysOfWeek[$date.DayOfWeek.value__] -eq "Sa" -or $daysOfWeek[$date.DayOfWeek.value__] -eq "Su") {
            $worksheet.Cells.Item($startRow +1 , $i + 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
            $weekendColumns += ($i + 1)
        }
        # I want to set the column width to 5 pixels for each day
        $worksheet.Columns.Item($i + 1).ColumnWidth = 5
    }
    $lastColumn = $worksheet.Cells.Item($startRow + 1, $daysInMonth + 1).Address(0, 0)
    SetFormulaHeaders $startRow $lastColumn
 
    
    # Create the drop-down list for the main data column
    for ($i = $startRow + 2; $i -le $startRow + $teamsize + 1; $i++) {
        for ($j = 2; $j -le $daysInMonth + 1; $j++) {
            $cell = $worksheet.Cells.Item($i, $j)
            $validation = $cell.Validation
            $validation.Delete()
            $validation.Add(3, 1, 1, $values)
            $validation.IgnoreBlank = $true
            $validation.InCellDropdown = $true

            # Check if the column index is in the $weekendColumns array
        if ($weekendColumns -contains $j) {
            $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)
        }
        }
        # Set the excel formulas after the headers are set hence the $startRow+2 
        SetExcelFormulas $i $lastColumn
    }
}

# Save the workbook
# Apply conditional formatting for "OS" cells
$range = $worksheet.UsedRange
$formatConditionOS = $range.FormatConditions.Add(1, 3, "OS")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkOrange)
$formatConditionOS.Font.Bold = $true

$formatConditionPTO = $range.FormatConditions.Add(1, 3, "PTO")  # xlCellValue = 1, xlEqual = 1
$formatConditionPTO.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Orange)
$formatConditionPTO.Font.Bold = $true

$formatConditionPTH = $range.FormatConditions.Add(1, 3, "PTH")  # xlCellValue = 1, xlEqual = 1
$formatConditionPTH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Silver)
$formatConditionPTH.Font.Bold = $true

$formatConditionOB = $range.FormatConditions.Add(1, 3, "OB")  # xlCellValue = 1, xlEqual = 1
$formatConditionOB.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Purple)
$formatConditionOB.Font.Bold = $true

$formatConditionOS = $range.FormatConditions.Add(1, 3, "APE")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Green)
$formatConditionOS.Font.Bold = $true

$formatConditionH = $range.FormatConditions.Add(1, 3, "H")  # xlCellValue = 1, xlEqual = 1
$formatConditionH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
$formatConditionH.Font.Bold = $true

$workbook.SaveAs($filePath)
$workbook.Close()
$excel.Quit()

# Release the COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "`nJob complete! Thank you!`n" -ForegroundColor Green