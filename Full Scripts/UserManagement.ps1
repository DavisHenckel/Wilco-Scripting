Function PrintMainMenu () {
    Clear-Host
    Write-Host ("============================================================================") -ForegroundColor Yellow
    Write-Host ("========================== WILCO USER MANAGEMENT ===========================") -ForegroundColor Yellow
    Write-Host ("============================ ") -NoNewline -ForegroundColor Yellow 
    Write-Host "Version 4.4.0" -ForegroundColor Cyan -NoNewline
    Write-Host (" =================================") -ForegroundColor Yellow
    Write-Host ("============================================================================`n") -ForegroundColor Yellow

    Write-Host "====================================================" -ForegroundColor Yellow
    Write-Host "Scripts Available                                  " -ForegroundColor Cyan -NoNewline
    Write-Host "|" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Yellow
    Write-Host "1: " -ForegroundColor Yellow -NoNewline
    Write-Host "New User" -ForegroundColor Magenta
    Write-Host "2: " -ForegroundColor Yellow -NoNewline
    Write-Host "Job or Location Change" -ForegroundColor Magenta
    Write-Host "3: " -ForegroundColor Yellow -NoNewline
    Write-Host "Separate User" -ForegroundColor Magenta
    Write-Host "4: " -ForegroundColor Yellow -NoNewline
    Write-Host "User Name Change" -ForegroundColor Magenta

    Write-Host "`n====================================================" -ForegroundColor Yellow
    Write-Host "Functions Available                                " -ForegroundColor Cyan -NoNewline
    Write-Host "|" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Yellow
    Write-Host "5: " -ForegroundColor Yellow -NoNewline
    Write-Host "Search for user by Employee ID" -ForegroundColor Magenta
    Write-Host "6: " -ForegroundColor Yellow -NoNewline
    Write-Host "Get All Users at Location" -ForegroundColor Magenta
    Write-Host "7: " -ForegroundColor Yellow -NoNewline
    Write-Host "Set User Password Changeable" -ForegroundColor Magenta
    Write-Host "8: " -ForegroundColor Yellow -NoNewline
    Write-Host "Get Users Password Expiration Info" -ForegroundColor Magenta
    Write-Host "9: " -ForegroundColor Yellow -NoNewline
    Write-Host "Show User in Global Address List" -ForegroundColor Magenta
    Write-Host "10: " -ForegroundColor Yellow -NoNewline
    Write-Host "Hide User in Global Address List" -ForegroundColor Magenta
    Write-Host "11: " -ForegroundColor Yellow -NoNewline
    Write-Host "Update Manager's `"Reports To`" to new Manager" -ForegroundColor Magenta
    Write-Host "12: " -ForegroundColor Yellow -NoNewline
    Write-Host "Sync Active Directory With Microsoft 365" -ForegroundColor Magenta

    Write-Host "`n====================================================" -ForegroundColor Yellow
    Write-Host "Reporting Tools                                    " -ForegroundColor Cyan -NoNewline
    Write-Host "|" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Yellow
    Write-Host "13: " -ForegroundColor Yellow -NoNewline
    Write-Host "Display Enabled Users with Office Licenses" -ForegroundColor Magenta
    Write-Host "14: " -ForegroundColor Yellow -NoNewline
    Write-Host "Display Disabled Users with Office Licenses" -ForegroundColor Magenta
    Write-Host "15: " -ForegroundColor Yellow -NoNewline
    Write-Host "Office Licenses assigned at all Location Codes" -ForegroundColor Magenta
    Write-Host "16: " -ForegroundColor Yellow -NoNewline
    Write-Host "Office Licenses assigned at one Location Code" -ForegroundColor Magenta
    Write-Host "17: " -ForegroundColor Yellow -NoNewline
    Write-Host "AD Audit Helper " -ForegroundColor Magenta
    Write-Host "18: " -ForegroundColor Yellow -NoNewline
    Write-Host "AD User Purge Helper " -ForegroundColor Magenta

    #Write-Host "`nEnter -Help to display documentation" -ForegroundColor Yellow
}

Function GetValidInput () {
    while ($true) {
        Write-Host "`n`nEnter your selection: " -ForegroundColor Cyan -NoNewline
        $UsersInput = Read-Host
        if ($UsersInput -eq "-exit") {
            return $null
        }
        try { [int32]$UsersInput = $UsersInput } #put in a try catch to validate integer
        catch { #if the user didn't enter an integer.
            Write-Host "Must enter a number..." -ForegroundColor Red
            Start-Sleep 2
            continue #prompt again
        }
        if ($UsersInput -gt 0 -and $UsersInput -lt 19) {
            return $UsersInput
        }
        else {
            Write-Host "Invalid input. Please enter an option 1-18." -ForegroundColor Red
            Start-Sleep 2
        }
    }
    return $UsersInput
}

Function InterpretInput ($Choice) {
    switch ($Choice) {
        1 {  
            Clear-Host
            $ScriptPath= $PSScriptRoot+"\InterpretNewHireEmail.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
        2 {
            Clear-Host
            $ScriptPath= $PSScriptRoot+"\InterpretJobChangeEmail.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
        3 {
            Clear-Host
            $ScriptPath= $PSScriptRoot+"\UserSeparation.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
        4 {
            Clear-Host
            $ScriptPath= $PSScriptRoot+"\NameChange.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
        5 {
            Clear-Host
            $EmployeeID = Read-Host -Prompt "Enter EmployeeID of user"
            $EmployeeID = ValidateEmployeeIDExists($EmployeeID)
            if ($null -eq $EmployeeID) {
                break
                $Script:PauseAfter = $true
            }
            try {
                $UserData = GetADUserCustom -SearchBy EmployeeID -SearchFor $EmployeeID -ReturnData All
            }
            catch {
                Write-Host "This user is disabled. Unable to retreive information about this user." -ForegroundColor Red
                $Script:PauseAfter = $true
                break
            }
            if ($null -eq $UserData) {
                Write-Host "This user is disabled. Unable to retreive information about this user." -ForegroundColor Red
                $Script:PauseAfter = $true
                break
            }
            $DeptName = GetDeptNameForLocation $UserData."Location"
            $OU = GetUsersCurrentOU $UserData."Distinguished Name"
            Write-Host -ForegroundColor Yellow -NoNewline "`nFull Name: "
            Write-Host -ForegroundColor Cyan $UserData."Name"
            Write-Host -ForegroundColor Yellow -NoNewline "sAMAccountName: "
            Write-Host -ForegroundColor Cyan $UserData."Logon Name"
            Write-Host -ForegroundColor Yellow -NoNewline "EmployeeID: "
            Write-Host -ForegroundColor Cyan $EmployeeID
            Write-Host -ForegroundColor Yellow -NoNewline "Location Info: "
            Write-Host -ForegroundColor Cyan "$($UserData."Location"), $DeptName"
            Write-Host -ForegroundColor Yellow -NoNewline "Job Title: "
            Write-Host -ForegroundColor Cyan $UserData."Job Title"
            Write-Host -ForegroundColor Yellow -NoNewline "Email Address: "
            Write-Host -ForegroundColor Cyan $UserData."Email Address"
            if ($UserData."Mobile Number") {
                Write-Host -ForegroundColor Yellow -NoNewline "Mobile Number: "
                Write-Host -ForegroundColor Cyan $UserData."Mobile Number"
            }
            Write-Host -ForegroundColor Yellow -NoNewline "OU: "
            Write-Host -ForegroundColor Cyan "$OU"
            Write-Host -ForegroundColor Yellow -NoNewline "Manager: "
            Write-Host -ForegroundColor Cyan "$($UserData."Manager")`n"
            $Script:PauseAfter = $true
        }
        6 {
            Clear-Host
            $LocationCode = Read-Host -Prompt "Enter the Location Code to lookup users by"
            $LocationCode = ValidateLocationCode $LocationCode
            $Users = GetAllUsersAtLocation $LocationCode
            Write-Host "`nUsers at location $LocationCode" -ForegroundColor Yellow
            Write-Host "=====================" -ForegroundColor Yellow
            ForEach ($User in $Users) {
                Write-host "$User" -ForegroundColor Cyan
            }
            $Script:PauseAfter = $true
        }
        7 {
            Clear-Host
            Write-Host -NoNewline -ForegroundColor Yellow "Enter the sAMAccountName of user to allow to change password: "
            $sAMAccountName = Read-Host
            SetPassChangeable $sAMAccountName
            $Script:PauseAfter = $true
        }
        8 {
            Clear-Host
            Write-Host -NoNewline -ForegroundColor Yellow "Enter the sAMAccountName of user to display password information: "
            $sAMAccountName = Read-Host
            GetPasswordExpiration $sAMAccountName
            $Script:PauseAfter = $true
        }
        9 {
            Clear-Host
            $EmployeeID = $null
            while($true) {
                try { [int32]$EmployeeID = Read-Host -Prompt "Enter EmployeeID of user to be shown in the Global Address List" }
                catch { 
                    Write-Host "Must enter a number." -ForegroundColor Red
                    continue
                }
                break #if we don't catch an error, break out of the loop.
            }
            $EmployeeID = ValidateEmployeeIDExists($EmployeeID)
            if ($null -eq $EmployeeID) {
                break
            }
            SetUserShownInGAL $EmployeeID
            $Script:PauseAfter = $true
        }
        10 { 
            Clear-Host
            $EmployeeID = $null
            while($true) {
                try { [int32]$EmployeeID = Read-Host -Prompt "Enter EmployeeID of user to be hidden in the Global Address List" }
                catch { 
                    Write-Host "Must enter a number." -ForegroundColor Red
                    continue
                }
                break #if we don't catch an error, break out of the loop.
            }
            $EmployeeID = ValidateEmployeeIDExists($EmployeeID)
            if ($null -eq $EmployeeID) {
                break
            }
            SetUserHiddenInGAL $EmployeeID
            $Script:PauseAfter = $true
        }
        11 {
            Clear-Host
            $ScriptPath= $PSScriptRoot+"\UpdateReportsTo.ps1"
            Invoke-Expression $ScriptPath
            $Script:PauseAfter = $false
        }
        12 {
            Clear-Host
            ADSYNCwithMS
            $Script:PauseAfter = $true
        }
        13 {
            Clear-Host
            DisplayUsersWithOfficeLicenses $true
            $Script:PauseAfter = $true

        }
        14 {
            Clear-Host
            DisplayUsersWithOfficeLicenses $false
            $Script:PauseAfter = $true
        }
        15 {
            Clear-Host
            $OfficeLicenseData = GetOfficeLicenseDataPerLocation
            WriteLicenseDataToExcel $OfficeLicenseData
            $Script:PauseAfter = $true
        }
        16 {
            Clear-Host
            Write-Host "Input location code to lookup Office License data at: " -ForegroundColor Yellow -NoNewline
            $LocationCode = Read-Host
            DisplayOfficeLicensesAtLocation $LocationCode
            $Script:PauseAfter = $true
        }
        17 {
            $ScriptPath= $PSScriptRoot+"\ADAuditHelp.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
        18 {
            $ScriptPath= $PSScriptRoot+"\ADUserPurge.ps1"
            Invoke-Expression $ScriptPath #jump to script
            $Script:PauseAfter = $false
        }
    }
}

#Main loop
SetupEnvironment
while ($true) {
    Clear-Variable -Name "PauseAfter" -ErrorAction SilentlyContinue
    PrintMainMenu
    $UsersChoice = GetValidInput
    if ($null -eq $UsersChoice) {
        break
    }
    InterpretInput $UsersChoice
    if ($Script:PauseAfter -eq $true) {
        Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
        [void][System.Console]::ReadKey($FALSE)
    }
}