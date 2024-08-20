
<#
.SYNOPSIS
Series of commands that helps build the CheckinApp to local server

.DESCRIPTION
  Performs building and running the CheckinApp. 

.PARAMETER year
  -$year -branchName parameter will pull and checkout desired branch.

.PARAMETER teamsize
  -$teamsize -Integer value of your teamsize e.g. 4 will create 4 rows for each team member. Default value is 1

.EXAMPLE 
    Displays help information for the script
        .\create.schedule.ps1 -? 
    
    Creates a schedule tracker for the year 2025 for default of 1 team member  
        .\create.schedule.ps1 -year 2025
    
    Creates a schedule tracker for the year 2026 for a team of 4 members
        ./create.schedule.ps1 -teamsize 4 -year 2026   
     

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
    [int]$teamsize = 1
)

Write-Host "`n`nCreating ScheduleTracker_$year.xlsx for a team of $teamsize.........." -ForegroundColor Blue

# Load the Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# Add a new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Define the different global variables
$daysOfWeek = @("Su", "M", "T", "W", "Th", "F", "Sa")
$values = "APE,H,OB,OS,PTO,PTH,WFA"
$columnMapping = @{
    29 = "AC"
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
    42 = "AP"
}
# Create a reverse mapping hashtable to map string values to integer representations
$reverseColumnMapping = @{}
foreach ($key in $columnMapping.Keys) {
    $reverseColumnMapping[$columnMapping[$key]] = $key
}
$weekdayColumnValue=$reverseColumnMapping["AG"]
$holidayColumnValue=$reverseColumnMapping["AH"]
$workingDaysColumnValue=$reverseColumnMapping["AI"]
$APEColumnValue = $reverseColumnMapping["AJ"]
$OBColumnValue = $reverseColumnMapping["AK"]
$OSColumnValue = $reverseColumnMapping["AL"]
$PTHColumnValue = $reverseColumnMapping["AM"]
$PTOColumnValue = $reverseColumnMapping["AN"]
$WFAColumnValue = $reverseColumnMapping["AO"]
$percentColumnValue = $reverseColumnMapping["AP"]
# Create an array to store the Row value of Names in the January Table
$arrayJanuaryNamesRows = @()
# Displays the Legend at the top of the sheet
$worksheet.Cells.Item(1, 1) = "APE"
$worksheet.Cells.Item(2, 1) = "H"
$worksheet.Cells.Item(3, 1) = "OS"
$worksheet.Cells.Item(4, 1) = "OB"
$worksheet.Cells.Item(1, 11) = "PTH"
$worksheet.Cells.Item(2, 11) = "PTO"
$worksheet.Cells.Item(3, 11) = "WFA"
$worksheet.Cells.Item(1, 2) = "Annual Physical Exam (0.5 Days by Default)"
$worksheet.Cells.Item(2, 2) = "Holiday"
$worksheet.Cells.Item(3, 2) = "Onsite"
$worksheet.Cells.Item(4, 2) = "Official Business (Business Trips, Client Visit, Conventions, Quarantine on OS Day, OS Day Canclled due to weather)"
$worksheet.Cells.Item(1, 12) = "Paid Time Off  - Half Day"
$worksheet.Cells.Item(2, 12) = "Paid Time Off (VL, SL, Maternity, Breavement)"
$worksheet.Cells.Item(3, 12) = "Work From Anywyare (PH Domestic/International Workcation)"

# Make column widths appropriate to the header text 
$worksheet.Columns.Item($reverseColumnMapping["AG"]).ColumnWidth = 8.9
$worksheet.Columns.Item($reverseColumnMapping["AI"]).ColumnWidth = 11.4
function SetFormulaHeaders ($startRow, $lastColumnHeading) {
    $nextRow = $startRow + 1
    $startRowPlus2 = $startRow + 2
    # Set the headers for columns with formulas 
    $worksheet.Cells.Item($startRow, $weekdayColumnValue) = "Weekdays"
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
    # Set the formula for weekdays
    $worksheet.Cells.Item($nextRow, $weekdayColumnValue).Formula = "=COUNTIF(B$($nextRow):$lastColumnHeading$($nextRow), `"M`") + COUNTIF(B$($nextRow):$lastColumnHeading$($nextRow), `"T`") + COUNTIF(B$($nextRow):$lastColumnHeading$($nextRow), `"W`") + COUNTIF(B$($nextRow):$lastColumnHeading$($nextRow), `"Th`") + COUNTIF(B$($nextRow):$lastColumnHeading$($nextRow), `"F`")"
    $worksheet.Cells.Item($startRow, $holidayColumnValue) = "Holidays"
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)  # Set font color
    # Set the formula for holidays
    $worksheet.Cells.Item($nextRow, $holidayColumnValue).Formula = "=COUNTIF(B$($startRowPlus2):AF$($startRowPlus2), `"H`")"
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue) = "Working Days"
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)  # Set background color
    # Set the formula for working days
    $worksheet.Cells.Item($nextRow, $workingDaysColumnValue).Formula = "=AG$nextRow - AH$nextRow"
     $worksheet.Cells.Item($nextRow, $APEColumnValue) = "APE"
     $worksheet.Cells.Item($nextRow, $OBColumnValue) = "OB"
    $worksheet.Cells.Item($nextRow, $OSColumnValue) = "OS"
    $worksheet.Cells.Item($nextRow, $PTHColumnValue) = "PTH"
    $worksheet.Cells.Item($nextRow, $PTOColumnValue) = "PTO"
    $worksheet.Cells.Item($nextRow, $WFAColumnValue) = "WFA"
    $worksheet.Cells.Item($nextRow, $percentColumnValue) = "%"
    $worksheet.Cells.Item($nextRow, $percentColumnValue).HorizontalAlignment = -4108  # Center alignment 
    $worksheet.Cells.Item($nextRow, $percentColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($nextRow, $percentColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkGreen)  # Set font color
}
function SetExcelFormulas ($startRow, $lastColumnHeading, $workingDaysRow) {
    $worksheet.Cells.Item($startRow, $APEColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$($startRow), `"APE`")/2"
    $worksheet.Cells.Item($startRow, $OBColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$($startRow), `"OB`")"
    $worksheet.Cells.Item($startRow, $OSColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$startRow, `"OS`")"    
    $worksheet.Cells.Item($startRow, $PTOColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$($startRow), `"PTO`")"
    $worksheet.Cells.Item($startRow, $PTHColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$($startRow), `"PTH`")/2"
    $worksheet.Cells.Item($startRow, $WFAColumnValue).Formula = "=COUNTIF(B$($startRow):$lastColumnHeading$($startRow), `"WFA`")"
    $worksheet.Cells.Item($startRow, $percentColumnValue).Formula = "=SUM(AJ$($startRow):AO$($startRow))/AI`$$($workingDaysRow)"
    $rangePercent = $worksheet.range("AP$($startRow)") # Set range of percentage column
    $rangePercent.NumberFormat = "0.0%"  # Set to % with 1 decimal place
}   

function SetBorders ($cellSetBorders) {
    # Set the border style for each cell
    $cellSetBorders.Borders.Item(9).LineStyle = 1 # xlEdgeBottom
    $cellSetBorders.Borders.Item(9).Weight = 2 # xlThin

    $cellSetBorders.Borders.Item(8).LineStyle = 1 # xlEdgeTop
    $cellSetBorders.Borders.Item(8).Weight = 2 # xlThin

    $cellSetBorders.Borders.Item(7).LineStyle = 1 # xlEdgeLeft
    $cellSetBorders.Borders.Item(7).Weight = 2 # xlThin

    $cellSetBorders.Borders.Item(10).LineStyle = 1 # xlEdgeRight
    $cellSetBorders.Borders.Item(10).Weight = 2 # xlThin
} 

# Get the current directory
$currentDirectory = Get-Location
$filePath= "$currentDirectory\ScheduleTracker_$year.xlsx"

# Delete existing file if it exists
if (Test-Path $filePath) {
    Remove-Item -Path $filePath
}

# Loop through each month of the year that was specified
for ($month = 1; $month -le 12; $month++) {
    
    
    $daysInMonth = [DateTime]::DaysInMonth($year, $month)
    $monthName = (Get-Date -Year $year -Month $month -Day 1).ToString("MMMM")
    Write-Host "Generating Table for $monthName" -ForegroundColor Cyan
    
    # Calculate the starting row for each month's table
    $startRow = ($month - 1) * ($teamsize + 2) + 5

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

    $lastColumn = $daysInMonth + 1
    SetFormulaHeaders $startRow $columnMapping[$lastColumn]
    $workingDaysRow= $startRow + 1
    
    $indexJanuaryNames = 0
    # Create the drop-down list for the main data column
    for ($i = $startRow + 2; $i -le $startRow + $teamsize + 1; $i++) {
        for ($j = 2; $j -le $daysInMonth + 1; $j++) {
            $cell = $worksheet.Cells.Item($i, $j)
            $validation = $cell.Validation
            $validation.Delete()
            $validation.Add(3, 1, 1, $values)
            $validation.IgnoreBlank = $true
            $validation.InCellDropdown = $true

            SetBorders $cell

            # Check if the column index is in the $weekendColumns array
            if ($weekendColumns -contains $j) {
                $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)
            }
        }
        if ($month -eq 1) {
            $arrayJanuaryNamesRows += $i
        }
        else {
            $worksheet.Cells.Item($i, 1) = "=A$($arrayJanuaryNamesRows[$indexJanuaryNames])"
            $indexJanuaryNames++ 
        }
        # Set the excel formulas after the headers are set and this is 3 rows from the start row hence the $i + 3
        SetExcelFormulas $i $columnMapping[$lastColumn] $workingDaysRow
    }
}

# Creats TOTAL table for MFA for each team member
$range = $worksheet.UsedRange
$currentLastRow = $range.Rows.Count
$lastRowInDecember = $currentLastRow
    $currentLastRow++
    $worksheet.Cells.Item($currentLastRow, 1).Value = "TOTAL"
    $currentLastRow++
    $worksheet.Cells.Item($currentLastRow, 1).Value = "WFA"
for ($i = 0; $i -lt $teamsize ; $i++) {
    $currentLastRow++
    $worksheet.Cells.Item($currentLastRow, 1).Value = "=A$($arrayJanuaryNamesRows[$i])"
    $worksheet.Cells.Item($currentLastRow, 2).Value = "=SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AF$($lastRowInDecember)=`"WFA`"))"
}

# Apply conditional formatting depending on cells values
$range = $worksheet.UsedRange

$formatConditionOS = $range.FormatConditions.Add(1, 3, "APE")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Green)
$formatConditionOS.Font.Bold = $true
$formatConditionH = $range.FormatConditions.Add(1, 3, "H")  # xlCellValue = 1, xlEqual = 1
$formatConditionH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
$formatConditionH.Font.Bold = $true
$formatConditionPTH = $range.FormatConditions.Add(1, 3, "PTH")  # xlCellValue = 1, xlEqual = 1
$formatConditionPTH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Silver)
$formatConditionPTH.Font.Bold = $true
$formatConditionPTO = $range.FormatConditions.Add(1, 3, "PTO")  # xlCellValue = 1, xlEqual = 1
$formatConditionPTO.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightBlue)
$formatConditionPTO.Font.Bold = $true
$formatConditionOB = $range.FormatConditions.Add(1, 3, "OB")  # xlCellValue = 1, xlEqual = 1
$formatConditionOB.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Purple)
$formatConditionOB.Font.Bold = $true
$formatConditionOS = $range.FormatConditions.Add(1, 3, "OS")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkOrange)
$formatConditionOS.Font.Bold = $true
$formatConditionOS = $range.FormatConditions.Add(1, 3, "WFA")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Blue)
$formatConditionOS.Font.Bold = $true
$formatConditionOS = $range.FormatConditions.Add(1, 3, "%")  # xlCellValue = 1, xlEqual = 1
$formatConditionOS.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkGreen)
$formatConditionOS.Font.Bold = $true
$rangePercent = $worksheet.range("AP2:AP$($worksheet.UsedRange.Rows.Count)") # Set range of percentage column starting from AO3
$rangePercent.NumberFormat = "0.0%"  # Set to % with 1 decimal place        
# Add conditional formatting for cells with values greater than or equal to 0.5
$formatConditionGreaterEqual50 = $rangePercent.FormatConditions.Add(1, 7, "0.5")  # xlCellValue = 1, xlGreaterEqual = 3
$formatConditionGreaterEqual50.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)
# Add conditional formatting for cells with values less than 0.5
$formatConditionLessThan50 = $rangePercent.FormatConditions.Add(1, 6, "0.5")  # xlCellValue = 1, xlLess = 2
$formatConditionLessThan50.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightPink)
$formatConditionTOTAL = $range.FormatConditions.Add(1, 3, "TOTAL")  # xlCellValue = 1, xlEqual = 1
$formatConditionTOTAL.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BlueViolet)
$formatConditionTOTAL.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)
$formatConditionTOTAL.Font.Bold = $true    

# Freeze pane at row 5, column 13 (M)
$worksheet.Application.ActiveWindow.SplitColumn = 21
$worksheet.Application.ActiveWindow.SplitRow = 4
$worksheet.Application.ActiveWindow.FreezePanes = $true
# Save the workbook
$workbook.SaveAs($filePath)
$workbook.Close()
$excel.Quit()

# Release the COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "`nJob complete! Thank you!`n" -ForegroundColor Green