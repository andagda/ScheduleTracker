
<#
.SYNOPSIS
Series of commands that creates an excel file for a team of employees to track their schedules for the year.

.DESCRIPTION
This script creates an excel file for a team of employees to track their schedules for the year. The script will prompt the user to enter the year and the number of employees in the team. The script will then generate an excel file with a table for each month of the year. Each table will have columns for each day of the month, and rows for each employee. The script will also calculate the number of weekdays, holidays, working days, and other types of days for each employee. The script will also calculate the percentage of working days for each employee.

.EXAMPLE 
    Displays help information for the script
        .\create.schedule.ps1 -? 
      
.NOTES
/*==================================================================================================
 = This file is part of the Navitaire CheckinApp application.
 = Copyright © Navitaire LLC, an Amadeus company. All rights reserved.
 =================================================================================================*/
#>

<# 
.PARAMETERS EXPLANATION:
year: Year that you want the schedule tracker to be created. This will create table from January to December When blank it will default to the current year.
#>

#region functions
function getData {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Schedule Tracker"
    $form.Size = New-Object System.Drawing.Size(300, 200)
    $form.StartPosition = "CenterScreen"

    $labelYear = New-Object System.Windows.Forms.Label
    $labelYear.Text = "Year:"
    $labelYear.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($labelYear)

    $textBoxYear = New-Object System.Windows.Forms.TextBox
    $textBoxYear.Location = New-Object System.Drawing.Point(110, 20)
    $form.Controls.Add($textBoxYear)

    $labelTeamSize = New-Object System.Windows.Forms.Label
    $labelTeamSize.Text = "Team Size:"
    $labelTeamSize.Location = New-Object System.Drawing.Point(10, 60)
    $form.Controls.Add($labelTeamSize)

    $textBoxTeamSize = New-Object System.Windows.Forms.TextBox
    $textBoxTeamSize.Location = New-Object System.Drawing.Point(150, 60)
    $form.Controls.Add($textBoxTeamSize)

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Location = New-Object System.Drawing.Point(50, 100)
    $buttonOK.Add_Click({
            if ($textBoxYear.Text -match '^\d{4}$' -and $textBoxTeamSize.Text -match '^\d+$') {
                $script:year = [int]$textBoxYear.Text
                $script:teamSize = [int]$textBoxTeamSize.Text
                $form.Close()
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Please enter valid values for Year and Number of Employees.")
            }
    })
    $form.Controls.Add($buttonOK)
    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Location = New-Object System.Drawing.Point(150, 100)
    $buttonCancel.Add_Click({ $form.Close() })
    $form.Controls.Add($buttonCancel)

    $form.ShowDialog()
}

function SetFormulaHeaders ($startRow, $lastColumnHeading) {
    $nextRow = $startRow + 1
    $startRowPlus2 = $startRow + 2
    # Set the headers for columns with formulas 
    $worksheet.Cells.Item($startRow, $weekdayColumnValue) = "Weekdays"
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $weekdayColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
    # Set the formula for weekdays
    $worksheet.Cells.Item($nextRow, $weekdayColumnValue).Formula = "=COUNTIF(D$($nextRow):$lastColumnHeading$($nextRow), `"M`") + COUNTIF(D$($nextRow):$lastColumnHeading$($nextRow), `"T`") + COUNTIF(D$($nextRow):$lastColumnHeading$($nextRow), `"W`") + COUNTIF(D$($nextRow):$lastColumnHeading$($nextRow), `"Th`") + COUNTIF(D$($nextRow):$lastColumnHeading$($nextRow), `"F`")"
    
    $worksheet.Cells.Item($startRow, $holidayColumnValue) = "Holidays"
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $holidayColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)  # Set font color
    # Set the formula for holidays
    $worksheet.Cells.Item($nextRow, $holidayColumnValue).Formula = "=COUNTIF(D$($startRowPlus2):AH$($startRowPlus2), `"H`")"

    $worksheet.Cells.Item($startRow, $workingDaysColumnValue) = "Working Days"
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($startRow, $workingDaysColumnValue).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)  # Set background color
    # Set the formula for working days
    $worksheet.Cells.Item($nextRow, $workingDaysColumnValue).Formula = "=AI$nextRow - AJ$nextRow"
    $worksheet.Cells.Item($nextRow, $WFAColumnValue) = "WFA"
    $worksheet.Cells.Item($nextRow, $OBColumnValue) = "OB"
    $worksheet.Cells.Item($nextRow, $WFOColumnValue) = "WFO"
    $worksheet.Cells.Item($nextRow, $PTHColumnValue) = "PTH"
    $worksheet.Cells.Item($nextRow, $PTOColumnValue) = "PTO"
    $worksheet.Cells.Item($nextRow, $percentColumnValue) = "%"
    $worksheet.Cells.Item($nextRow, $percentColumnValue).HorizontalAlignment = -4108  # Center alignment 
    $worksheet.Cells.Item($nextRow, $percentColumnValue).Font.Bold = $true
    $worksheet.Cells.Item($nextRow, $percentColumnValue).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
}
function SetExcelFormulas ($startRow, $lastColumnHeading, $workingDaysRow) {
    $worksheet.Cells.Item($startRow, $WFAColumnValue).Formula = "=`COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"WFA`") + (COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"WFA-H`")/2)"
    $worksheet.Cells.Item($startRow, $OBColumnValue).Formula = "=COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"OB`")"
    $worksheet.Cells.Item($startRow, $WFOColumnValue).Formula = "=COUNTIF(D$($startRow):$lastColumnHeading$startRow, `"WFO`") + (COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"WFO-H`")/2)"    
    $worksheet.Cells.Item($startRow, $PTOColumnValue).Formula = "=COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"PTO`")"
    $worksheet.Cells.Item($startRow, $PTHColumnValue).Formula = "=COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"PTH`")/2 + (COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"WFA-H`")/2) + (COUNTIF(D$($startRow):$lastColumnHeading$($startRow), `"WFO-H`")/2)"
    $worksheet.Cells.Item($startRow, $workingDaysPerEmployeeColumnValue).Formula = "=AK`$$($workingDaysRow) - SUM(AO$($startRow):AQ$($startRow))"
    $worksheet.Cells.Item($startRow, $percentColumnValue).Formula = "=SUM(AL$($startRow):AN$($startRow))/(B$($startRow))"
    $rangePercent = $worksheet.range("C$($startRow)") # Set range of percentage column
    $rangePercent.NumberFormat = "0.0%"  # Set to % with 1 decimal place
    # Add conditional formatting for cells with values greater than or equal to 0.5
    $formatConditionGreaterEqual50 = $rangePercent.FormatConditions.Add(1, 7, "0.5")  # xlCellValue = 1, xlGreaterEqual = 3
    $formatConditionGreaterEqual50.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)
    # Add conditional formatting for cells with values less than 0.5
    $formatConditionLessThan50 = $rangePercent.FormatConditions.Add(1, 6, "0.5")  # xlCellValue = 1, xlLess = 2
    $formatConditionLessThan50.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightPink)
    
    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionLessThan50) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionGreaterEqual50) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($rangePercent) | Out-Null
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
#endregion functions

getData



# Use the collected values
if ($year -and $teamSize) {
    Write-Host "Year: $year"
    Write-Host "Number of Employees: $teamSize"
    # Call your existing script logic here with the collected values
    # .\create.schedule.ps1 -year $year -teamsize $teamSize
    Write-Host "`n`nCreating ScheduleTracker_$year.xlsx for a team of $teamsize.........." -ForegroundColor Blue

    # Load the Excel COM object
    $excel = $null
    $workbook = $null
    $worksheet = $null
    $excelProcessId = $null
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Capture the Excel process ID for cleanup verification
        $excelProcessId = (Get-Process | Where-Object { $_.MainWindowHandle -eq $excel.Hwnd }).Id

        # Add a new workbook
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

    # Define the different global variables
    $daysOfWeek = @("Su", "M", "T", "W", "Th", "F", "Sa")
    $values = "WFA,WFA-H,H,OB,WFO,WFO-H,PTO,PTH"
    
    # Load holidays from JSON file
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $holidaysJsonPath = Join-Path $scriptPath "holidays.json"
    $holidaysData = $null
    $regularHolidays = @()
    $floatingHolidays = @()
    
    if (Test-Path $holidaysJsonPath) {
        try {
            $holidaysJson = Get-Content -Path $holidaysJsonPath -Raw | ConvertFrom-Json
            $regularHolidays = $holidaysJson.regularHolidays
            $floatingHolidays = $holidaysJson.floatingHolidays
            Write-Host "Loaded holidays from holidays.json" -ForegroundColor Green
        }
        catch {
            Write-Host "Warning: Could not load holidays.json. Holidays will not be auto-populated." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Warning: holidays.json not found. Holidays will not be auto-populated." -ForegroundColor Yellow
    }
    
    $columnMapping = @{
        2 = "B"
        3 = "C"
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
        43 = "AQ"
    }
    # Create a reverse mapping hashtable to map string values to integer representations
    $reverseColumnMapping = @{}
    foreach ($key in $columnMapping.Keys) {
        $reverseColumnMapping[$columnMapping[$key]] = $key
    }
    $workingDaysPerEmployeeColumnValue = $reverseColumnMapping["B"]
    $percentColumnValue = $reverseColumnMapping["C"]
    $weekdayColumnValue = $reverseColumnMapping["AI"]
    $holidayColumnValue = $reverseColumnMapping["AJ"]
    $workingDaysColumnValue = $reverseColumnMapping["AK"]
    $WFAColumnValue = $reverseColumnMapping["AL"]
    $OBColumnValue = $reverseColumnMapping["AM"]
    $WFOColumnValue = $reverseColumnMapping["AN"]
    $PTHColumnValue = $reverseColumnMapping["AO"]
    $PTOColumnValue = $reverseColumnMapping["AP"] 
    
    # Create an array to store the Row value of Names in the January Table
    $arrayJanuaryNamesRows = @()
    # Displays the Legend at the top of the sheet
    $worksheet.Cells.Item(1, 1) = "H"
    $worksheet.Cells.Item(2, 1) = "WFO"
    $worksheet.Cells.Item(3, 1) = "WFA"
    $worksheet.Cells.Item(4, 1) = "OB"
    $worksheet.Cells.Item(1, 10) = "PTH"
    $worksheet.Cells.Item(2, 10) = "PTO"
    $worksheet.Cells.Item(1, 22) = "WFO-H"
    $worksheet.Cells.Item(2, 22) = "WFA-H"
    $worksheet.Cells.Item(1, 2) = "Holiday"
    $worksheet.Cells.Item(2, 2) = "Work From Office"
    $worksheet.Cells.Item(4, 2) = "Official Business (Business Trips, Client Visit, Conventions, Quarantine on WFO Day, WFO Day Cancelled due to weather)"
    $worksheet.Cells.Item(1, 11) = "Paid Time Off  - Half Day (APE, VL, SL, Maternity, Bereavement)"
    $worksheet.Cells.Item(2, 11) = "Paid Time Off (APE, VL, SL, Maternity, Bereavement)"
    $worksheet.Cells.Item(3, 2) = "Work From Anywhere (PH Domestic/International Workcation)"
    $worksheet.Cells.Item(1, 23) = "Work from Office with Half Day PTO"
    $worksheet.Cells.Item(2, 23) = "Work from Anywhere with Half Day PTO"

    # Make column widths appropriate to the header text 
    $worksheet.Columns.Item($reverseColumnMapping["AI"]).ColumnWidth = 8.9
    $worksheet.Columns.Item($reverseColumnMapping["AK"]).ColumnWidth = 11.4

    # Get the current directory
    $currentDirectory = Get-Location
    $filePath = "$currentDirectory\ScheduleTracker_$year.xlsx"

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
        $worksheet.Cells.Item($startRow + 1, 1).Value = "Name"
        $worksheet.Cells.Item($startRow + 1, 1).HorizontalAlignment = -4108  # Center alignment
        $worksheet.Cells.Item($startRow + 1, 1).Font.Bold = $true
        $worksheet.Cells.Item($startRow + 1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Green)  # Set background color
        $worksheet.Cells.Item($startRow + 1, 1).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
        $worksheet.Cells.Item($startRow + 1, 2).Value = "Working Days"
        $worksheet.Cells.Item($startRow + 1, 2).HorizontalAlignment = -4108  # Center alignment
        $worksheet.Cells.Item($startRow + 1, 2).Font.Bold = $true
        $worksheet.Cells.Item($startRow + 1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkCyan)  # Set background color
        $worksheet.Cells.Item($startRow + 1, 2).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
        $worksheet.Cells.Item($startRow + 1, 3).Value = "%"
        $worksheet.Cells.Item($startRow + 1, 3).HorizontalAlignment = -4108  # Center alignment
        $worksheet.Cells.Item($startRow + 1, 3).Font.Bold = $true
        $worksheet.Cells.Item($startRow + 1, 3).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkMagenta)  # Set background color
        $worksheet.Cells.Item($startRow + 1, 3).Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)  # Set font color
    
        # hash table to store the weekend columns
        $weekendColumns = @()
        $fillerInt = 3
        # Set the main headers (dates of the month)
        for ($i = 1; $i -le $daysInMonth; $i++) {
            $date = Get-Date -Year $year -Month $month -Day $i
            $worksheet.Cells.Item($startRow, $i + $fillerInt) = $date.ToString("dd")
            $worksheet.Cells.Item($startRow + 1, $i + $fillerInt) = $daysOfWeek[$date.DayOfWeek.value__]
            if ($daysOfWeek[$date.DayOfWeek.value__] -eq "Sa" -or $daysOfWeek[$date.DayOfWeek.value__] -eq "Su") {
                $worksheet.Cells.Item($startRow + 1 , $i + $fillerInt).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)  # Set background color
                $weekendColumns += ($i + 3)
            }
            # I want to set the column width to 5 pixels for each day
            $worksheet.Columns.Item($i + 1).ColumnWidth = 5
        }

        $lastColumn = $daysInMonth + $fillerInt
        SetFormulaHeaders $startRow $columnMapping[$lastColumn]
        $workingDaysRow = $startRow + 1
    
        $indexJanuaryNames = 0
        # Create the drop-down list for the main data column
        for ($i = $startRow + 2; $i -le $startRow + $teamsize + 1; $i++) {
            for ($j = 4; $j -le $daysInMonth + $fillerInt; $j++) {
                $cell = $worksheet.Cells.Item($i, $j)
                $validation = $cell.Validation
                $validation.Delete()
                $validation.Add(3, 1, 1, $values)
                $validation.IgnoreBlank = $true
                $validation.InCellDropdown = $true

                # Check if this day is a holiday and falls on a weekday
                $dayOfMonth = $j - $fillerInt
                $currentDate = Get-Date -Year $year -Month $month -Day $dayOfMonth
                $dayOfWeekValue = $currentDate.DayOfWeek.value__
                
                # Check if it's a weekday (Monday=1 to Friday=5)
                $isWeekday = ($dayOfWeekValue -ge 1 -and $dayOfWeekValue -le 5)
                
                if ($isWeekday) {
                    # Check regular holidays
                    $isHoliday = $false
                    foreach ($holiday in $regularHolidays) {
                        if ($holiday.month -eq $month -and $holiday.day -eq $dayOfMonth) {
                            $cell.Value = "H"
                            $isHoliday = $true
                            break
                        }
                    }
                    
                    # Check floating holidays if not already a regular holiday
                    if (-not $isHoliday) {
                        foreach ($holiday in $floatingHolidays) {
                            if ($holiday.month -eq $month -and $holiday.day -eq $dayOfMonth) {
                                $cell.Value = "H"
                                break
                            }
                        }
                    }
                }

                SetBorders $cell

                # Check if the column index is in the $weekendColumns array
                if ($weekendColumns -contains $j) {
                    $cell.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGray)
                }
                
                # Release COM objects
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($validation) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null
            }
            if ($month -eq 1) {
                $arrayJanuaryNamesRows += $i
            }
            else {
                $worksheet.Cells.Item($i, 1) = "=A$($arrayJanuaryNamesRows[$indexJanuaryNames])"
                $indexJanuaryNames++ 
            }
            SetExcelFormulas $i $columnMapping[$lastColumn] $workingDaysRow
        }
    }

    # Creates TOTAL table for WFA for each team member
    $range = $worksheet.UsedRange
    $currentLastRow = $range.Rows.Count
    $lastRowInDecember = $currentLastRow
    $currentLastRow++
    $worksheet.Cells.Item($currentLastRow, 1).Value = "TOTAL"
    $currentLastRow++
    $worksheet.Cells.Item($currentLastRow, 2).Value = "WFA"
    $worksheet.Cells.Item($currentLastRow, 3).Value = "PTO"
    for ($i = 0; $i -lt $teamsize ; $i++) {
        $currentLastRow++
        $worksheet.Cells.Item($currentLastRow, 1).Value = "=A$($arrayJanuaryNamesRows[$i])"
        $worksheet.Cells.Item($currentLastRow, 2).Value = "=SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"WFA`")) + SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"WFA-H`"))" # WFA-H is counted as full WFA day
        $worksheet.Cells.Item($currentLastRow, 3).Value = "=SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"PTO`")) + (SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"PTH`"))/2) + (SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"WFA-H`"))/2) + (SUMPRODUCT((A$($arrayJanuaryNamesRows[0]):A$($lastRowInDecember)=A$($currentLastRow))*(B$($arrayJanuaryNamesRows[0]):AH$($lastRowInDecember)=`"WFO-H`"))/2)"
    }

    # Apply conditional formatting depending on cells values
    $range = $worksheet.UsedRange

    $formatConditionH = $range.FormatConditions.Add(1, 3, "H")  # xlCellValue = 1, xlEqual = 1
    $formatConditionH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Red)
    $formatConditionH.Font.Bold = $true
    $formatConditionPTH = $range.FormatConditions.Add(1, 3, "PTH")  # xlCellValue = 1, xlEqual = 1
    $formatConditionPTH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkGreen)
    $formatConditionPTH.Font.Bold = $true
    $formatConditionPTO = $range.FormatConditions.Add(1, 3, "PTO")  # xlCellValue = 1, xlEqual = 1
    $formatConditionPTO.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Green)
    $formatConditionPTO.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightYellow)
    $formatConditionPTO.Font.Bold = $true
    $formatConditionOB = $range.FormatConditions.Add(1, 3, "OB")  # xlCellValue = 1, xlEqual = 1
    $formatConditionOB.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Purple)
    $formatConditionOB.Font.Bold = $true
    $formatConditionWFO = $range.FormatConditions.Add(1, 3, "WFO")  # xlCellValue = 1, xlEqual = 1
    $formatConditionWFO.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkOrange)
    $formatConditionWFO.Font.Bold = $true
    $formatConditionWFOH = $range.FormatConditions.Add(1, 3, "WFO-H")  # xlCellValue = 1, xlEqual = 1
    $formatConditionWFOH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Orange)
    $formatConditionWFOH.Font.Bold = $true
    $formatConditionWFA = $range.FormatConditions.Add(1, 3, "WFA")  # xlCellValue = 1, xlEqual = 1
    $formatConditionWFA.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Blue)
    $formatConditionWFA.Font.Bold = $true
    $formatConditionWFAH = $range.FormatConditions.Add(1, 3, "WFA-H")  # xlCellValue = 1, xlEqual = 1
    $formatConditionWFAH.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::DarkBlue)
    $formatConditionWFAH.Font.Bold = $true     
    $formatConditionTOTAL = $range.FormatConditions.Add(1, 3, "TOTAL")  # xlCellValue = 1, xlEqual = 1
    $formatConditionTOTAL.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BlueViolet)
    $formatConditionTOTAL.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)
    $formatConditionTOTAL.Font.Bold = $true    

    # Release format condition COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionTOTAL) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionWFAH) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionWFA) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionWFOH) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionWFO) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionOB) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionPTO) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionPTH) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($formatConditionH) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null

    # Freeze pane at row 5, column 3 (M)
    $worksheet.Application.ActiveWindow.SplitColumn = 3
    $worksheet.Application.ActiveWindow.SplitRow = 4
    $worksheet.Application.ActiveWindow.FreezePanes = $true
    $worksheet.Columns.Item(2).ColumnWidth = 12
    $worksheet.Columns.Item(33).ColumnWidth = 5
    $worksheet.Columns.Item(34).ColumnWidth = 5
    # Save the workbook
    $workbook.SaveAs($filePath)
    
        Write-Host "`nJob complete! Thank you!`n" -ForegroundColor Green
    }
    catch {
        Write-Host "`nError occurred: $_" -ForegroundColor Red
    }
    finally {
        # Release the COM objects properly
        if ($workbook) {
            try { $workbook.Close($false) } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch { }
            $workbook = $null
        }
        if ($worksheet) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null } catch { }
            $worksheet = $null
        }
        if ($excel) {
            try { 
                $excel.DisplayAlerts = $false
                $excel.Quit() 
            } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
            $excel = $null
        }
        
        # Force garbage collection multiple times to ensure COM objects are released
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Force kill the Excel process if it's still running
        if ($excelProcessId) {
            Start-Sleep -Milliseconds 500
            $excelProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            if ($excelProcess) {
                try {
                    $excelProcess | Stop-Process -Force
                    Write-Host "Forcefully closed remaining Excel process." -ForegroundColor Yellow
                } catch {
                    Write-Host "Note: Excel process may still be running. Check Task Manager if needed." -ForegroundColor Yellow
                }
            }
        }
    }
}

