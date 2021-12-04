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