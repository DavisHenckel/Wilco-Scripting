#Ensure elevated prompt
Function SetupEnvironment {
    $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

    # Get the security principal for the Administrator role
    $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

    # Check to see if we are currently running "as Administrator"
    if ($myWindowsPrincipal.IsInRole($adminRole)) {
        # We are running "as Administrator" - so change the title and background color to indicate this
        $Host.UI.RawUI.WindowTitle = "Wilco User Management"
        $Host.UI.RawUI.BackgroundColor = "Black"
        Clear-Host
    }
    else {
        Write-Host "Please Re-Run this Script as Administrator"
        cmd /c pause
        exit
    }
}

#Takes in an employee number
#Removes all groups from employee
Function SetUserToHaveNoGroups($EmpNum) {
    Try {
        $username = Get-ADUser -Filter {employeeid -eq $EmpNum} | Select-Object name,sAMAccountName,UserPrincipalName | Sort-Object name
    }
    Catch {
        Write-Host "Unable to find  `"$EmpNum`" in AD."-ForegroundColor Red
        return
    }
    Get-ADUser -Identity $($username.sAMAccountName) -Properties MemberOf | ForEach-Object {$_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}
}

#Takes in an arraylist of groups
#Takes in an employee number
#Assigns all groups to employee number
Function SetUserToHaveGroups($GroupsToAdd, $EmpNum) {
    Try {
        $Username = Get-ADUser -Filter {employeeid -eq $EmpNum} | Select-Object name,sAMAccountName,UserPrincipalName
    }
    Catch {
        Write-Host "Unable to locate `"$EmpNum`" in AD." -ForegroundColor Red
        return
    }
    if ($null -eq $Username) {
        Write-Host "User `"$EmpNum`" does not exist in AD.`nRe-run script with correct employee number." -ForegroundColor Red
    }
    else {
        if ($GroupsToAdd.length -eq 0) {
            Write-Host "SetUserToHaveGroups got passed an empty group list. The user will not be added to any additional groups." -ForegroundColor Yellow
            return
        }
        ForEach ($Group in $GroupsToAdd) { #assigns the groups to the user.
            if ($Group -match "MB: ") {
                continue
            }
            Try {
                Add-ADGroupMember -Identity $Group -Members $Username.sAMAccountName
            }
            Catch { #if this fails, group name is likely wrong or the group doesn't exist in local AD.
                Write-Host "Unable to add $($Username.name) to AD group: $Group" -ForegroundColor Red
                Write-Host "Trying online groups now..." -ForegroundColor Yellow
                Start-Sleep 2
                AddUserToOnlineSG $EmpNum $Group
            }
        }
    }
}

#Slightly modified function for new user scripts
#Takes in an arraylist of groups
#Takes in an employee number
#Takes in an office license assignment as well.
#Assigns all groups to employee number
#TODO Refactor at some point
Function SetUserToHaveGroups_NewUserScript($GroupsToAdd, $EmpNum, $OfficeAssigned) {
    Try {
        $Username = Get-ADUser -Filter {employeeid -eq $EmpNum} | Select-Object name,sAMAccountName,UserPrincipalName
    }
    Catch {
        Write-Host "Unable to locate `"$EmpNum`" in AD." -ForegroundColor Red
        return
    }
    if ($null -eq $Username) {
        Write-Host "User `"$EmpNum`" does not exist in AD.`nRe-run script with correct employee number." -ForegroundColor Red
    }
    else {
        if ($GroupsToAdd.length -eq 0) {
            Write-Host "SetUserToHaveGroups got passed an empty group list. The user will not be added to any additional groups." -ForegroundColor Yellow
            return
        }
        ForEach ($Group in $GroupsToAdd) { #assigns the groups to the user.
            if ($Group -match "MB: ") {
                # Write-Host "Cannot add $Group yet. This will be added as an update to the script in the future." -ForegroundColor Red
                continue
            }
            Try {
                Add-ADGroupMember -Identity $Group -Members $Username.sAMAccountName
            }
            Catch { #if this fails, group name is likely wrong or the group doesn't exist in local AD.
                Write-Host "Unable to add $($Username.name) to AD group: $Group" -ForegroundColor Red
                if ($OfficeAssigned -eq $true) {
                    Write-Host "Trying online groups now..." -ForegroundColor Yellow
                    Start-Sleep 2
                    AddUserToOnlineSG $EmpNum $Group
                }
                else {
                    Write-Host "Will not assign online groups without Office license assignment. You will need to process manually."
                }
            }
        }
    }
}

#Takes in an OU as a string
#Takes in an employee number
#Places the employee in the OU.
Function SetUserInOU($OUToBePlacedIn, $EmpNum) {
    $Username = Get-ADUser -Filter {employeeid -eq $EmpNum} | Select-Object name,sAMAccountName,UserPrincipalName
    if ($null -eq $Username) {
        Write-Host "User `"$EmpNum`" does not exist in AD.`nRe-run script with correct employee number." -ForegroundColor Red
    }
    Get-ADUser -Filter {employeeID -eq $EmpNum} | Move-ADObject -TargetPath $OUToBePlacedIn
}

#Takes in an employee number
#Sets the msExchHideFromAddressList as true.
Function SetUserHiddenInGAL ($EmployeeID) {
    $EmployeeID = ValidateEmployeeIDExists $EmployeeID
    Try {
        $Username = Get-ADUser -Filter {employeeid -eq $EmployeeID} | Select-Object name,sAMAccountName | Sort-Object nameWrite-
    }
    Catch {
        Write-Host "ERROR -- Unable to locate `"$EmployeeID`" in the AD. Try again." -ForegroundColor Red
    }
    Try{
        Write-Host "Attempting to hide user in GAL..." -ForegroundColor Magenta
        Set-ADUser -Identity $($Username.sAMAccountName) -Replace @{msExchHideFromAddressLists=$true}
        $var = $($Username.sAMAccountName)
    }
    Catch {
        Write-Host "ERROR -- Unable to hide `"$var`" in the GAL. Try again." -ForegroundColor Red
    }
    Write-host "`"$var`" has been hidden from Global Address List" -ForegroundColor Green
}

#Takes in an employee number
#Sets the msExchHideFromAddressList as false.
Function SetUserShownInGAL ($EmployeeID) {
    $EmployeeID = ValidateEmployeeIDExists $EmployeeID
    $Username = Get-ADUser -Filter {employeeid -eq $EmployeeID} | Select-Object name,sAMAccountName | Sort-Object name
    Try {
        Write-Host "Attempting to show user in GAL..." -ForegroundColor Magenta
        Set-ADUser -Identity $($Username.sAMAccountName) -replace @{msExchHideFromAddressLists=$false}
        $var = $($Username.sAMAccountName)
    }
    Catch {
        Write-Host "ERROR -- Unable to show user `"$var`" in the GAL. Try again." -ForegroundColor Red
        return
    }
    Write-Host "User `"$var`" is now shown in the Global Address List" -ForegroundColor Green
}

#Takes in a distinguished name of manager
#Takes in a SAMAccountName of a user
#Sets the distName to be SAMNames Manager
Function SetUserToHaveNewManager ($ManagerDistName, $UserToChangeSAMName) {
    Try {
        Set-ADUser -Identity $UserToChangeSAMName -Replace @{manager=$ManagerDistName} #modify the attribute
        Write-Host "Successfully updated $UserToChangeSAMName to have $ManagerDistName as their manager!" -ForegroundColor Green
    }
    Catch {
        Write-Host "Unable to update $UserToChangeSAMName to have $ManagerDistName as their manager" -ForegroundColor Red
    }
}

# Takes in a manager distinguishedname
# Takes in an arraylist of employees SAM account names.
# Sets all users in the arraylist to have the manager dist name.
Function SetAllUsersToHaveNewManager ($ManagerDistName, $ArrayListOfEmployees) {
    if ($ArrayListOfEmployees.length -eq 0) {
        Write-Host "There are no users to update." -ForegroundColor Yellow
        return
    }
    $AreYouSure = 'n'
    Write-Host "Users to Update" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Yellow
    ForEach ($User in $ArrayListOfEmployees) {
        Write-Host $User -ForegroundColor Yellow
    }
    Write-Host "`nARE YOU SURE YOU WISH TO UPDATE THE ABOVE USERS TO HAVE " -ForegroundColor Magenta
    Write-Host -ForegroundColor Cyan $ManagerDistName
    Write-Host "AS THEIR MANAGER? (y/n): "-NoNewline -ForegroundColor Magenta
    While($true) {
        $AreYouSure = Read-Host
        Write-Host ""
        if ($AreYouSure -eq "n" -or $AreYouSure-eq "N" -or $AreYouSure -eq "y" -or $AreYouSure -eq "Y") {
            break
        }
        Write-Host "Invalid input. Enter y/n. "
    }
    if ($AreYouSure -eq "y" -or $AreYouSure -eq "Y") {
        Write-Host "Updating users to have this manager..." -ForegroundColor Magenta
        ForEach ($User in $ArrayListOfEmployees) {
            Write-Host "User " -NoNewline -ForegroundColor Magenta
            Write-Host "$User" -ForegroundColor Cyan -NoNewline
            Write-Host " now has manager " -ForegroundColor Magenta -NoNewline
            Write-Host "$ManagerDistName" -ForegroundColor Cyan
            Set-ADUser -Identity $User -Replace @{manager=$ManagerDistName} #modify the attribute
        }
    }
    else {
        Write-Host "Not updating any users' managers."
    }
    return
}

#Takes in a SAMAccountNAme 
#Sets the password changeable attribute so the user can change their password prior to 24hrs.
Function SetPassChangeable ($sAMAccountName) {
    Try {
        Set-ADUser -Identity $sAMAccountName -Replace @{pwdLastSet=0}
    }
    Catch {
        Write-Host "Unable to find user `"$sAMAccountName`"" -ForegroundColor Red
        return
    }
    Write-Host "Attribute pwdLastSet modified successfully. User should be able to change password now." -ForegroundColor Green
    return
}

#Takes in a location code
#Takes in an employee ID
#Sets the user's AD Attributes to have that location 
Function SetUserLocation ($LocationCode, $EmployeeID) {
    $LocationCode = ValidateLocationCode($LocationCode) #ensure Location code is valid
    $EmployeeID = ValidateEmployeeIDExists($EmployeeID) #ensure user Id is valid.
    $DataToAssign = GetLocationInfo $LocationCode #Loads the row from Location Reference Spreadsheet
    $Username = Get-ADUser -Filter {employeeid -eq $EmployeeID} | Select-Object name, sAMAccountName
    
    ForEach ($Key in ($DataToAssign.GetEnumerator())) { #loop through all elements of the row for the location code in location reference spreadsheet.
        if ($Key.Value -match '^\s$') { #regex if value is only whitespace, replace with null
            $Key.Value= $null #if empty string, assign to null so attribute is not set in AD, rather than " "
        }
        if ($Key.Name -eq "Domain" -or $Key.Name -eq "OU") { 
            continue #we don't assign these
        }
        if ($Key.Name -eq "Manager") {
            $Key.Value = GetManagerAtLocation $LocationCode #get manager at location
        }
        Try { #try this first, there are 2 ways of assigning AD attributes.
            Set-ADUser -Identity $($Username.sAMAccountName) -Replace @{$Key.Name=$Key.Value}
        }
        Catch { #try this next if it doesn't work. If for some reason this fails, a powershell error will occur.
            #Write-Host "Catch -- $($Key.Name)=$($Key.Value)" #print statements for debugging
            $PSDefaultParameterValues=@{"Set-ADUser:$($Key.Name)"=$Key.Value} #see https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_parameters_default_values?view=powershell-7.1
            Set-ADUser -Identity $($Username.sAMAccountName) #Set ADUSer with the set DefaultParameterValues which are set on the line above
        }
        $PSDefaultParameterValues.Clear() #clear hash for safety.
    }
}