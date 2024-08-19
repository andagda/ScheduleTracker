## create.schedule.ps1 tool to Create a New Schedule Tracker Sheet

1. Run the script the following from a powershell command line 
 `.\create.schedule.ps1 -?`
 or 
 `./create.schedule.ps1 -?` to know how to use the script.
1. To get all the Details from the help file use `Get-Help .\create.schedule.ps1 -Detailed` 
1. To get the Examples from the help file details use `Get-Help .\create.schedule.ps1 -Examples`
1. Here is what should be displayed on your powershell CLI when running the script with these parameters `.\create.schedule.ps1 -teamsize 4 -year 2026`.  
![image.jpg](help.images\CLIscreenshot.jpg)
1. An Excel File with **ScheduleTracker_YYYY.xlsx** will be generated in the root folder  
![image.jpg](help.images\ExcelFileSample.jpg)
1. Creates a Table for Each Month of the specified year  
![image.jpg](help.images\ExcelFileMonthsTable.jpg)
1. All formulas needed is pre-populated for each month  
Weekday Formula  
![image.jpg](help.images\ExcelFilesFormulaWeekdays.jpg)  
OS or Onsite Formula  
![image.jpg](help.images\ExcelFilesFormulaOS.jpg)
RTO % Formula  
![image.jpg](help.images\ExcelFileFormulaRTOPercent.jpg)


 
