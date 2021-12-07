$LicenseNames = @{ #this is a hash to store the office licenses respective names.
    "DOMAINSPECIFICPREFIX:ENTERPRISEPACK" = "E3";
    "DOMAINSPECIFICPREFIX:EXCHANGEDESKLESS" = "Exchange Online Kiosk";
    "DOMAINSPECIFICPREFIX:POWER_BI_PRO" = "Power BI Pro"; 
    "DOMAINSPECIFICPREFIX:DESKLESSPACK" = "F3";
    "DOMAINSPECIFICPREFIX:STANDARDPACK" = "E1";
    "DOMAINSPECIFICPREFIX:FLOW_FREE" = "Flow Free";
    "DOMAINSPECIFICPREFIX:VISIOCLIENT" = "Visio Plan 2";
    "DOMAINSPECIFICPREFIX:POWER_BI_STANDARD" = "Power BI Free";
    "DOMAINSPECIFICPREFIX:MICROSOFT_BUSINESS_CENTER" = "Microsoft Business Center";
    "DOMAINSPECIFICPREFIX:STREAM" = "Microsoft Stream Trial";
    "DOMAINSPECIFICPREFIX:PROJECTPROFESSIONAL" = "Project Plan 3";
    "DOMAINSPECIFICPREFIX:TEAMS_EXPLORATORY" = "Microsoft Teams Exploratory";
    "DOMAINSPECIFICPREFIX:POWERAUTOMATE_ATTENDED_RPA" = "Power Automate per user with attended RPA Plan"
}

#Function to sync AD with MS Online
Function ADSYNCwithMS {
    Write-Host "Syncing Active Directory with Microsoft 365..." -ForegroundColor Magenta
    $PotentialError1 = $null
    Try {
        Invoke-Command -ComputerName ComputerSpecificName -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -ErrorAction SilentlyContinue -ErrorVariable $PotentialError1 | Out-Null
    }
    Catch {
        Write-Host "ERROR -- Unable to sync with MS Online. Run the command again." -ForegroundColor Red
        return
    }
    Write-Host "Sync completed successfully" -ForegroundColor Green
}

#Takes in a UPN (emaill addr)
#Returns boolean on whether the user is shown in MSOL environment so licenses can be assigned.
Function IsUserShownInMSOL ($UserPrincipalName) {
    $PotentialError = $null
    Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorVariable $PotentialError -ErrorAction SilentlyContinue
    if ($PotentialError -ne "") {
        return $false
    }
    return $true
}

#Takes in nothing
#Checks to ensure MSOL Module is connected and credentials aren't needed
Function ConnectMSOL {
    Get-MsolDomain -ErrorAction SilentlyContinue | Out-Null
    if (-not $?) { #if MSOL is not connected already, it likely expired.
        Connect-MsolService
        return
    }
    return
}

#Takes in nothing
#Checks to ensure EXO Module is connected and credentials aren't needed
Function ConnectEXO {
    $getsessions = Get-PSSession | Select-Object -Property State, Name
    $isconnected = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
    if ($isconnected -ne "True") {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowBanner:$false
    }
}

#Takes in nothing
#Checks to ensure Azure AD Module is connected and credentials aren't needed
Function ConnectAzureAD {
    Try {
        $temp = Get-AzureADTenantDetail | Out-Null #this won't work unless AzureAD is connected. Assign to vaiable and Out-Null to prevent output to console
    }
    Catch {
        Connect-AzureAD #if not connected, connect
    }
    if ($temp) { #just to get rid of dumb VScode warning about unused variable.
        return
    }
}

#Takes in a username
#Returns boolean on whether the user is shown in EXO environment so mailboxes can be assigned.
Function IsUserShownInEXO ($Name) {
    $PotentialError = $null
    Get-User -Identity $Name -ErrorVariable $PotentialError -ErrorAction SilentlyContinue
    if ($PotentialError -ne "") {
        return $false
    }
    return $true
}

#Takes in an employee ID
#Takes in a group name
#Adds the user to the online security group, if the group couldn't be added, it calls the function AddUserToOnlineDG
Function AddUserToOnlineSG ($EmployeeID, $OnlineGroup) {
    ConnectMSOL #Ensure MSOL is connected. 
    $GroupID = Get-Msolgroup | Where-Object {$_.displayName -eq $OnlineGroup -and $null -eq $_.lastdirsynctime}  -ErrorAction SilentlyContinue | Select-Object ObjectID #captures the group ID 
    $OnlineOnlySGs = GetAllOnlineOnlySG
    if ($OnlineOnlySGs -contains $OnlineGroup -eq $false) {
        AddUserToOnlineDG $EmployeeID $OnlineGroup
        return
    }
    $UsersData = GetADUserCustom -SearchBy EmployeeID -SearchFor $EmployeeID -ReturnData All
    $Name = $UsersData."Name"
    $UserPrincipalName = $UsersData."Email Address"
    $UserExistsOnline = IsUserShownInMSOL $UserPrincipalName
    $Counter = 0
    while ($UserExistsOnline -eq $false) {
        if ($Counter -eq 0) {
            Write-Host "Waiting for user to be shown in MSOL.`nThis may take a while for new users." -ForegroundColor Magenta -NoNewline
        }
        $UserExistsOnline = IsUserShownInMSOL $UserPrincipalName
        $Counter = $Counter + 1
        if ($Counter % 20 -eq 0) { #output a . every 20 tries to show the user it's still working.
            Write-Host -NoNewline "." -ForegroundColor Magenta
        }
    }
    if ($Counter -gt 0) {
        Write-Host ""
    }
    $UserGUID = Get-MsolUser -UserPrincipalName $UserPrincipalName | Select-Object ObjectID #captures the user GUID
    Try {
        Add-MsolGroupMember -GroupObjectId $GroupID."ObjectId" -GroupMemberType User -GroupMemberObjectId $UserGUID."ObjectId" -ErrorAction Stop #Add the user to the group
        Write-Host "Added `"" -ForegroundColor Green -NoNewline
        Write-Host $Name -ForegroundColor Magenta -NoNewline
        Write-Host "`" to group `""-ForegroundColor Green -NoNewline
        Write-Host $OnlineGroup -ForegroundColor Magenta -NoNewline
        Write-Host "`" as a Security Group successfully!" -ForegroundColor Green
        Start-Sleep 2
    }
    Catch {
        AddUserToOnlineDG $EmployeeID $OnlineGroup
        return
    }
}

#Takes in an employee ID
#Takes in an online Group
#Assigns the user to the online group, outputs any errors if there were errors found.
Function AddUserToOnlineDG ($EmployeeID, $OnlineGroup) {
    ConnectEXO #Ensure Exchange Online is connected.
    $GroupName = Get-DistributionGroup -Identity $OnlineGroup -ErrorAction SilentlyContinue | Select-Object Name 
    if ($null -eq $GroupName) {
        Write-Host "Unable to locate $OnlineGroup as an online Group. You will need to assign manually.`n" -ForegroundColor Red
        return
    }
    $ToAdd = $GroupName."Name"
    $Name = GetADUserCustom -SearchBy EmployeeID -SearchFor $EmployeeID -ReturnData Name
    $UserExistsOnline = IsUserShownInEXO $Name
    $Counter = 0
    while ($UserExistsOnline -eq $false) {
        if ($Counter -eq 1000) {
        Write-Host "Unable to add user `"$Name`" to Online Distribution or Security Group: `"$OnlineGroup`"" -ForegroundColor Red
        return
        }
        if ($Counter -eq 750) {
            Write-Host "`nCurrently searching for user: $Name and it is taking a long time. Try entering another name to search for: " -ForegroundColor Yellow -NoNewline
            $Name = Read-Host
        }
        if ($Counter -eq 0) {
            Write-Host "Waiting for user to be shown in Exchange Online.`nThis may take a while for new users." -ForegroundColor Magenta -NoNewline
        }
        $UserExistsOnline = IsUserShownInEXO $Name
        $Counter = $Counter + 1
        if ($Counter % 20 -eq 0) { #output a . every 20 tries to show the user it's still working.
            Write-Host -NoNewline "." -ForegroundColor Magenta
        }
    }
    if ($Counter -gt 0) {
        Write-Host ""
    }
    $UserPrincipalName = GetADUserCustom -SearchBy Name -SearchFor $Name -ReturnData "Mail"
    Try {
        Add-DistributionGroupMember -Identity $ToAdd -Member $UserPrincipalName -ErrorAction Stop #Add the user to the group
        Write-Host "Added `"" -ForegroundColor Green -NoNewline
        Write-Host $Name -ForegroundColor Magenta -NoNewline
        Write-Host "`" to group `""-ForegroundColor Green -NoNewline
        Write-Host $ToAdd -ForegroundColor Magenta -NoNewline
        Write-Host "`" as a Distribution Group successfully!" -ForegroundColor Green
        Start-Sleep 1
    }
    Catch {
        Write-Host "Unable to add user `"$Name`" to Online Distribution or Security Group: `"$OnlineGroup`"" -ForegroundColor Red
        return
    }
}

#Takes in a SAMACcountName
#Signs user out of all Sessions
Function SignUserOutOfM365 ($sAMAccountName) {
    ConnectAzureAD
    $UsersEmail = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $err1 = $null
    Get-AzureADUser -SearchString $UsersEmail | Revoke-AzureADUserAllRefreshToken -ErrorAction SilentlyContinue -ErrorVariable $err1 
    if ($null -ne $err1) {
        Write-Host "Failed to sign user out of all sessions" -ForegroundColor Red
        return
    }
    Write-Host "Successfully signed user out of all sessions." -ForegroundColor Green
}

#Takes in Nothing
#Returns an arraylist containing all the online only distribution groups
Function GetAllOnlineOnlyDG {
    ConnectMSOL
    $Groups = Get-Msolgroup -All | Where-Object {$null -eq $_.lastdirsynctime -and $_."GroupType" -eq "Distribution"} | Select-Object DisplayName
    $DGNames = [System.Collections.ArrayList]@()
    ForEach($Group in $Groups) {
        $DGNames.Add($Group."DisplayName") | Out-Null
    }
    return $DGNames
}

#Takes in Nothing
#Returns an arraylist containing all the online only security groups
Function GetAllOnlineOnlySG {
    ConnectMSOL
    $Groups = Get-Msolgroup | Where-Object {$null -eq $_.lastdirsynctime -and $_."GroupType" -eq "Security"} | Select-Object DisplayName
    $SGNames = [System.Collections.ArrayList]@()
    ForEach($Group in $Groups) {
        $SGNames.Add($Group."DisplayName") | Out-Null
    }
    return $SGNames
}

#https://social.technet.microsoft.com/Forums/exchange/en-US/8f8a4aaa-a6c1-424c-886f-5ea69ef7e328/remove-a-user-from-all-the-distribution-group?forum=exchangesvradminlegacy
#Takes in an arraylist of groups to be removed
#Takes in a SAMAccountName
#Removes the user from all the groups, if errors were encountered, it will return the number of errors.
#Returns number of errors encountered
Function RemoveUserFromOnlineDGs ($GroupsToRemove, $sAMAccountName) {
    ConnectEXO
    $ErrorsEncountered = $null
    $UsersEmail = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    if ($GroupsToRemove.Length -lt 1) {
        return
    }
    ForEach ($DG in $GroupsToRemove) {
        $GrpName = Get-AzureADGroup -ObjectId $DG | Select-Object DisplayName #get group name instead of ObjectID
        try {
            Remove-DistributionGroupMember -Confirm:$false -Identity $DG -member $UsersEmail -ErrorAction Stop
            Write-Host "Removed " -ForegroundColor Green -NoNewline
            Write-Host $UsersEmail -ForegroundColor Cyan -NoNewline
            Write-Host " from group "-ForegroundColor Green -NoNewline
            Write-Host $GrpName."DisplayName" -ForegroundColor Cyan -NoNewline
            Write-Host " successfully!" -ForegroundColor Green
        }
        catch {
            try {
                Remove-UnifiedGroupLinks -Identity $DG -LinkType Members -Links $UsersEmail -Confirm:$False -ErrorAction Stop
                Write-Host "Removed " -ForegroundColor Green -NoNewline
                Write-Host $UsersEmail -ForegroundColor Cyan -NoNewline
                Write-Host " from group "-ForegroundColor Green -NoNewline
                Write-Host $GrpName."DisplayName" -ForegroundColor Cyan -NoNewline
                Write-Host " successfully!" -ForegroundColor Green
            }
            catch {
                $ErrorsEncountered = 1
                Write-Host "Failed to remove " -ForegroundColor Red -NoNewline
                Write-Host $UsersEmail -ForegroundColor Cyan -NoNewline
                Write-Host " from group "-ForegroundColor Red -NoNewline
                Write-Host $GrpName."DisplayName" -ForegroundColor Cyan
            }
        }
    }
    return $ErrorsEncountered
}

#https://social.technet.microsoft.com/Forums/exchange/en-US/8f8a4aaa-a6c1-424c-886f-5ea69ef7e328/remove-a-user-from-all-the-distribution-group?forum=exchangesvradminlegacy
#Takes in an arraylist of groups to be removed
#Takes in a SAMAccountName
#Removes the user from all the groups, if errors were encountered, it will return the number of errors.
#Returns number of errors encountered
Function RemoveUserFromOnlineSGs ($GroupsToRemove, $sAMAccountName) {
    ConnectMSOL
    $ErrorsEncountered = $null
    $UsersEmail = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    if ($GroupsToRemove.Length -lt 1) {
        return
    }
    ForEach ($OnlineGroup in $GroupsToRemove) {
        $GrpName = Get-AzureADGroup -ObjectId $OnlineGroup | Select-Object DisplayName #get group name instead of ObjectID
        $err1 = $null
        $UserGUID = Get-MsolUser -UserPrincipalName $UsersEmail | Select-Object ObjectID #captures the user GUID
        Remove-MsolGroupMember -GroupObjectId $OnlineGroup -GroupMemberType User -GroupmemberObjectId $UserGUID."ObjectId" -ErrorVariable $err1 -ErrorAction SilentlyContinue
        if ($null -ne $err1) {
            Write-Host "Failed to reomve `"" -ForegroundColor Green -NoNewline
            Write-Host $UsersEmail -ForegroundColor Cyan -NoNewline
            Write-Host "`" from Online Security group `""-ForegroundColor Green -NoNewline
            Write-Host $GrpName."DisplayName" -ForegroundColor Cyan 
            $ErrorsEncountered = 1
            $ErrorsEncountered = 1
            continue
        }
        else{
            Write-Host "Removed " -ForegroundColor Green -NoNewline
            Write-Host $UsersEmail -ForegroundColor Cyan -NoNewline
            Write-Host " from group "-ForegroundColor Green -NoNewline
            Write-Host $GrpName."DisplayName" -ForegroundColor Cyan -NoNewline
            Write-Host " successfully!" -ForegroundColor Green
            Start-Sleep 1
            continue
        }
    }
    Return $ErrorsEncountered
}

#Takes in a SAMAccountName
#Removes the user from all online only groups.
Function RemoveUserFromOnlineGroups ($sAMAccountName) {
    ConnectAzureAD #maintain persistent connection
    $UserEmail = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $User = Get-AzureADUser -ObjectId $UserEmail
    $UserMembership = Get-AzureADUserMembership -ObjectId $User."ObjectID" #Get the groups the user is a part of.
    $DGGroupsToRemove = [System.Collections.ArrayList]@() #Group to store 
    $SGGroupsToRemove = [System.Collections.ArrayList]@() #Group to store 
    $SGs = GetAllOnlineOnlySG
    $DGs = GetAllOnlineOnlyDG
    ForEach ($Group in $UserMembership) { #loop through users groups
        if ($SGs.Contains($Group."DisplayName")) { #if the group is an online only group
            $SGGroupsToRemove.Add($Group."ObjectId") | Out-Null #add GUID to array
        }
        if ($DGs.Contains($Group."DisplayName")) {
            $DGGroupsToRemove.Add($Group."ObjectId") | Out-Null #add GUID to array
        }
    }
    $Errors1 = RemoveUserFromOnlineSGs $SGGroupsToRemove $sAMAccountName
    $Errors2 = RemoveUserFromOnlineDGs $DGGroupsToRemove $sAMAccountName
    if ($Errors1) { #if not null from function
        Write-Host "Encountered errors when removing user from Online Online Security Groups`nVerify removal of Groups in the Admin Center." -ForegroundColor Red
    }
    if ($Errors2) { #if not null from function
        Write-Host "Encountered errors when removing user from Online Online Distribution Groups`nVerify removal of Groups in the Admin Center." -ForegroundColor Red
    }
}

#https://community.spiceworks.com/topic/1982283-office-365-remove-all-licenses-from-a-user
#Takes in a SAMAccountName
#Removes the office license for that user.
Function RemoveOfficeLicensesFromUser ($sAMAccountName) {
    ConnectMSOL
    $EmailAddress = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    (Get-MsolUser -UserPrincipalName $EmailAddress).licenses.AccountSkuId | ForEach-Object {
        Write-Host "Removing Office 365 License: " -NoNewline -ForegroundColor Magenta
        $LicenseName = $LicenseNames."$_"
        Write-Host "$LicenseName" -ForegroundColor Cyan -NoNewline
        Write-Host " from " -ForegroundColor Magenta -NoNewline
        Write-Host "$EmailAddress" -ForegroundColor Cyan -NoNewline
        Write-Host "..." -ForegroundColor Magenta
        Set-MsolUserLicense -UserPrincipalName $EmailAddress -RemoveLicenses $_
    }
}

#Takes in a boolean
#Returns to the console, the users that are enabled or disabled, depending on the first argument.
Function DisplayUsersWithOfficeLicenses ($Enabled) {
    ConnectMSOL
    $ListOfUsers = $null
    if ($Enabled -eq $true) {
        Write-Host "Enabled Users who have Office Licenses" -ForegroundColor Yellow
        Write-Host "======================================" -ForegroundColor Magenta
        $ListOfUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | Select-Object UserPrincipalName #Get the list of all users that are enabled in our O365 environment
    }
    elseif ($Enabled -eq $false) {
        Write-Host "Disabled Users who have Office Licenses" -ForegroundColor Yellow
        Write-Host "=======================================" -ForegroundColor Magenta
        $ListOfUsers = Get-MsolUser -All -EnabledFilter DisabledOnly | Select-Object UserPrincipalName #Get the list of all users that are enabled in our O365 environment
    }
    else {
        Write-Host "Please Pass a boolean to indicate whether searching for enabled or disabled users."
        return
    }
    ForEach ($User in $ListOfUsers) { #loop through all users
        $Upn = $User."UserPrincipalName"
        (Get-MsolUser -UserPrincipalName $Upn).licenses.AccountSkuId | ForEach-Object { #Get all the user's office licenses
            if ($null -eq $_) { #if no office license, we don't care. Skip to next user.
                continue
            }
            $License = $LicenseNames."$_" #Load License by key $_
            Write-Host "$Upn : $License" -ForegroundColor Cyan
        }
    }
    Write-Host "Function completed." -ForegroundColor Green
    return
}