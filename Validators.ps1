#Takes in a location code
#Returns a validated location code.
Function ValidateLocationCode ($LocationCode) {
    while ($true) {
        $LocationCode = "$LocationCode"
        $PathToFile = $PSScriptRoot + "\Excel Dependencies\FileName.xlsx"
        $ExcelFile = $null
        Try {
            $ExcelFile = Import-Excel $PathToFile -WorksheetName "LocationInfo"
        }
        Catch {
            Write-Host "ERROR -- Could not read file at path $($PathToFile). Exiting" -ForegroundColor Red
            cmd /c pause
            exit
        }
        $ArrayOfLocCodes = $ExcelFile."departmentNumber"
        $Count = 0
        ForEach ($LocCode in $ArrayOfLocCodes) {  #Convert to text since spreadsheed downloads as numbers.
            $ArrayOfLocCodes[$Count] = "$LocCode"
            $Count = $Count + 1
        }
        $IndexOfLocationCode = $ArrayOfLocCodes.IndexOf($LocationCode)
        if ($IndexOfLocationCode -eq -1) {
            Write-Host "ERROR -- Unable to find Location Code `"$LocationCode`"" -ForegroundColor Red
            Write-Host -NoNewline -ForegroundColor Yellow "Enter a valid location code. Enter -Exit to return to the previous menu: "
            $LocationCode = Read-Host
            if ($LocationCode -eq "-exit" -or $LocationCode -eq "-Exit") {
                return $null
            }
        }
        else {
            break
        }
    }
    return $LocationCode
}

#Imports the excel sheet with all the valid Job titles.
#Takes in a job title and returns a valid job title.
Function ValidateJobTitle ($JobTitle) {
    Try {
        $ExcelFilePath = $PSScriptRoot + "\Excel Dependencies\FileName.xlsm"
        $AccessFile = Import-Excel $ExcelFilePath -WorksheetName "WorkSheetName"
    }
    Catch {
        Write-Host "ERROR -- Could not read file. Exiting" -ForegroundColor Red
        Pause
        exit
    }
    $JobTitleArray = $AccessFile."Job Titles"
    while ($true) { #infinite loop
        if ($JobTitleArray.Contains($JobTitle)) {
            break
        }
        #Allows an invalid job title to be entered.
        if ($JobTitle -match " -OVERRIDE") {
            return $JobTitle
        }
        Write-Host "`nERROR -- Job Title `"$JobTitle`" doesn't exist.`nYou can override this validation and continue the script with an unvalidated Job title if you like.`nTo do so, enter the job title followed by `" -OVERRIDE`"`nExample: SomeJobNotDefined -OVERRIDE" -ForegroundColor Red
        $JobTitle = Read-Host -Prompt "Enter a Job Title"
        Write-Host "`n"
    }
    return $JobTitle
}

# Takes in an EmployeeID and checks in AD if the Employee ID exists. 
# Returns an existing employee's ID
Function ValidateEmployeeIDExists ($EmployeeID) {
    while ($true) {
        if ($EmployeeID -match "^\d+$" -eq 0 -or $EmployeeID.length -gt 5) { #checks for length and ensures it is numeric and less than 6 chars
            Write-Host ("`nERROR, Employee ID `"$EmployeeID`" contains non numeric characters or is greater than 5 characters`n") -ForegroundColor Red
            $EmployeeID = Read-Host -Prompt ("Enter the correct Employee ID")
            continue
        }
        $UserTest = Get-ADUser -Filter {employeeID -eq $EmployeeID}
        if ($null -ne $UserTest) { #If user does exist in AD
            return $EmployeeID
        }
        else { #if user does not exist in AD
            Write-Host "User `"$EmployeeID`" does not exist in AD." -ForegroundColor Red
            Write-Host -NoNewline -ForegroundColor Yellow "Enter an existing Employee ID. Enter -Exit to go back to the previous menu: "
            $EmployeeID = Read-Host
            if ($EmployeeID -eq "-exit" -or $EmployeeID -eq "-Exit") {
                return $null
            }
            Write-Host "`n"
            continue
        }
    }
}

# Takes in an EmployeeID and checks in AD if the Employee ID is unused and is available. 
# Returns an unused employee ID
Function ValidateEmployeeIDAvailable ($EmployeeID) {
    while ($true) {
        if ($EmployeeID -match "^\d+$" -eq 0 -or $EmployeeID.length -gt 5) { #checks for length and ensures it is numeric and less than 6 chars
            Write-Host ("`nERROR, Employee ID `"$EmployeeID`" contains non numeric characters or is greater than 5 characters`n") -ForegroundColor Red
            $EmployeeID = Read-Host -Prompt ("Enter the correct Employee ID")
            continue
        }
        $UserTest = Get-ADUser -Filter {employeeID -eq $EmployeeID}
        if ($null -ne $UserTest) { #If user does exist in AD
            Write-Host "User `"$EmployeeID`" already exists in AD." -ForegroundColor Red
            Write-Host -NoNewline -ForegroundColor Yellow "Enter a valid Employee ID. Enter -Exit to return to the previous menu: "
            $EmployeeID = Read-Host
            Write-Host "`n"
            if ($EmployeeID -eq "-Exit" -or $EmployeeID -eq "-exit") {
                return $null
            }
            continue
        }
        else { #if user does not exist in AD
            return $EmployeeID
        }
    }
}

