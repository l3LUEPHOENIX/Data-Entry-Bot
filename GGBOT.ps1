### Data Entry Bot ###
# Input data taken from chat logs into excel form.
# All Files involved with the transaction:
# - GGBOT.ps1 <--- The script itself.
# - counter_yyyyMMdd.txt <--- Where the line count value is stored.
# - details.txt <--- Details that may be sensitive in nature live here.
# - place1_yyyyMMdd.txt <--- The chat log file being monitored.
# - GGBOT_LOGS_yyyyMMdd.txt <--- The log file storing a history of all bot activities.
# - log.xlsx <--- Where data is being inputted to. [NOT IMPLEMENTED YET]

$getdate = Get-Date -Format "yyyyMMdd"

# If it's a new day, and a log file doesn't exist, create the log and it's corresponding counter file
if ((Test-Path ".\place1_$(Get-Date -Format "yyyyMMdd").txt") -and (-not (Test-Path ".\GGBOT_LOGS_$getdate.txt"))) { # Test to see if new chat log file has been created. If not, create the file and provide the header.
    Write-Output "########## GGBOT LOG FOR $getdate ##########" > "GGBOT_LOGS_$getdate.txt"
    Write-Output "0" > "counter_$getdate.txt"
}

##### STORED DATA #####
# This keeps sensitive information out of the code!
$details = Get-Content .\details.txt
$alpha = $details[0].Split(";")[1] 
$bravo = $details[1].Split(";") 
$charlie = $details[2].Split(";") 
$delta = $details[3].Split(";") 
$echo = $details[4].Split(";") 
$foxtrot = $details[5].Split(";") 

$currentRoom = ".\place2_$getdate.txt" # The value stored in this variable is based upon the naming convention of the generated log files.
$entry = Get-Content $currentRoom
$oldItems = [int](Get-Content ".\counter_$getdate.txt") # Stored in counter.txt is a single integer, which is the number of counted lines since the last update
$newItems = $entry.Length # This gets the new line count from the file.

if ($oldItems -lt $newItems) { # Test if the number of previously counted lines differs from the current line count.
    Foreach ($line in $oldItems..$newItems) { 
        if ($entry[$line].Contains($bravo[1]) -or $entry[$line].Contains($bravo[2]) -and (-not $entry[$line].Split(" ")[1].Contains("$alpha"))) { 
            
            ### Init Variables ###
            $newEntry = $entry[$line] 
            $row = 1

            ### The function for all excel inputs will live here. ###
            # Open the workbook
            $excelEntry = New-Object -comobject Excel.Application
            $openExcel = $excelEntry.Workbooks.Open("C:\Path\To\Document\log.xlsx")
            $selectSheet = $openExcel.Sheets.Item(1) # This needs to check date and return abbreviation of the month

            # Go into the sheet and add to cells. 
            while ($selectSheet.Cells.Item($row,1).Value2 -notlike "") { $row++ }
            $selectSheet.Cells.Item($row,1).value2 = "RED"
            $selectSheet.Cells.Item($row,3).value2 = "$newEntry"
            $selectSheet.Cells.Item($row,4).value2 = "GG BOT" 
            $gamma = $newEntry.Split(" ")[4]
            $bravoCol = ""
            switch ("($gamma)") {
                "($($echo[1]))" {$bravoCol += $foxtrot[1]}
                "($($echo[2]))" {$bravoCol += $foxtrot[2]}
                default {"N/A"}
            }
            $bravoCol += " ($($newEntry.Split(" ")[3])) $($newEntry.Split(" ")[2])"
            $selectSheet.Cells.Item($row,2).value2 = "$bravoCol"
            $selectSheet.Cells.Item($row,2).WrapText = "True"
            
            # Save and close out excel.
            $openExcel.Save()
            $excelEntry.Quit()

            ### Write to Bot Log ###
            Write-Output "[$(Get-Date -Format "HH:mm")][FROM LINE: $line TO ROW: $row][$newEntry]" >> "GGBOT_LOGS_$getdate.txt" # Add the new entry to the log and add a time stamp and cell number
            Write-Host "[$(Get-Date -Format "HH:mm")][CELL] > New Entry from line: $line to row: $row"

        }
     }
     Write-Output "$newItems" > ".\counter_$getdate.txt" # Update the counter to the current line count
}

### NOTES BELOW ### 
# Will need to generate a new counter file per room OR overwrite the counter file back to zero. [DONE]
# Encapsulate this into a tick function that can be passed as a job; I only want to click "Run Once" once....
# Still needs to select sheet by getting the date and formatting it to a three letter abreviation.
