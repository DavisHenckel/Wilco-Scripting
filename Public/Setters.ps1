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