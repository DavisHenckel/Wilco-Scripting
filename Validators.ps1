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