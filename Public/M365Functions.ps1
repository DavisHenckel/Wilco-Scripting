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

#Takes in a SAMAccountName
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

#Takes in a license name.
# Returns true if the given Office 365 License is Available and can be assigned
# Returns false if the given license is not available
Function IsO365LicenseAvailable ($LicenseName) {
    ConnectMSOL
    $O365License = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -eq $LicenseName}
    $NumAvailableLicenses = $O365License.ActiveUnits - $O365License.ConsumedUnits
    If ($NumAvailableLicenses -lt 1) {
        return $false
    }
    return $true
}

#Takes in a SAMACcountName
#Takes in an office license to be added
#Assigns the office license to the user.
Function AssignOfficeLicense ($sAMAccountName, $Office365LicenseToAdd) {
    $LicName = $Office365LicenseToAdd
    $EmailAddress = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    ConnectMSOL
    $PotentialError1 = $null
    $PotentialError2 = $null
    Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation US -ErrorVariable PotentialError1 -ErrorAction SilentlyContinue
    Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $Office365LicenseToAdd -ErrorVariable PotentialError2 -ErrorAction SilentlyContinue
    if ($PotentialError1 -ne "" -or $PotentialError2 -ne "") {
        #Write-Host "ERROR -- Unable to assign $Office365LicenseToAdd to $EmailAddress." -ForegroundColor Red
        return $false
    }
    Write-Host "Added " -ForegroundColor Green -NoNewline
    Write-Host "$LicName" -ForegroundColor Cyan -NoNewline
    Write-Host " license successfully!" -ForegroundColor Green 
    return $true
}

#Takes in a SAMACcountName
#Takes in an office license to be added
#Assigns the office license to the user.
Function AssignOfficeLicense ($sAMAccountName, $Office365LicenseToAdd) {
    $LicName = $Office365LicenseToAdd
    $EmailAddress = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    ConnectMSOL
    $PotentialError1 = $null
    $PotentialError2 = $null
    Set-MsolUser -UserPrincipalName $EmailAddress -UsageLocation US -ErrorVariable PotentialError1 -ErrorAction SilentlyContinue
    Set-MsolUserLicense -UserPrincipalName $EmailAddress -AddLicenses $Office365LicenseToAdd -ErrorVariable PotentialError2 -ErrorAction SilentlyContinue
    if ($PotentialError1 -ne "" -or $PotentialError2 -ne "") {
        #Write-Host "ERROR -- Unable to assign $Office365LicenseToAdd to $EmailAddress." -ForegroundColor Red
        return $false
    }
    Write-Host "Added " -ForegroundColor Green -NoNewline
    Write-Host "$LicName" -ForegroundColor Cyan -NoNewline
    Write-Host " license successfully!" -ForegroundColor Green 
    return $true
}

#Takes in nothing
#Returns a HashTable containing all location codes and the count of each office license they have.
Function GetOfficeLicenseDataPerLocation {
    Write-Host "Office License Per Location Report." -ForegroundColor Yellow
    Write-Host "===================================" -ForegroundColor Yellow
    ConnectMSOL
    $SavePath = $env:USERPROFILE + "\Desktop\UnknownLocationUsers.csv"
    Write-Host "If location cannot be determined, the user and their license will be written to UnknownLocationUsers.csv at:`n" -ForegroundColor Yellow
    Write-Host "    $SavePath`n" -ForegroundColor Cyan
    Write-Host "Collecting Office License information. This will take a while...(approximately 8 minutes)" -ForegroundColor Magenta
    if (Test-Path $SavePath) {
        Remove-Item -Force $SavePath
    }
    # Write-Output "This file displays errors in Office License lookup. Likely because of no location code, or the user was not located in AD. All of these are placed in the `"Unknown`" Section of the displayed data." | Out-File -Append -FilePath "C:\ErrorsInOfficeLicenses.txt"
    $LicensesPerLoc = @{ } #Hash we will store data in. This starts as a 1d hash but turns into a 2d hash
    $ListOfUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | Select-Object UserPrincipalName #Get the list of all users that are enabled in our O365 environment
    ForEach ($User in $ListOfUsers) { #loop through all users
        $Upn = $User."UserPrincipalName"
        if ($null -eq $Upn) {
            continue
        }
        (Get-MsolUser -UserPrincipalName $Upn).licenses.AccountSkuId | ForEach-Object { #Get all the user's office licenses
            if ($null -eq $_) { #if no office license, we don't care. Skip to next user.
                continue
            }
            else {
                $sAMAccountName = GetADUserCustom -SearchBy Mail -SearchFor $Upn -ReturnData "sAMAccountName" #if this user has a license, get their AD UPN. It should be the same, but sometimes it's not. 
                $License = $LicenseNames."$_" #load the Office licenses name rather than SKU.
                if (-not $sAMAccountName -and $License) {
                    $OutputObj = @{
                        "User" = $Upn;
                        "License" = $License;
                        "Reason for Unknown" = "AD Account Issue, (maybe name or UPN?)"
                    }
                    $OutputObject = [PSCustomObject]$OutputObj
                    $OutputObject | Export-Csv -Path $SavePath -Append -NoTypeInformation
                    $LocCode = "Unknown"
                    $CurrentCountOfLicenses = $LicensesPerLoc.$LocCode.$License #load current number of this license per this location
                    if (-not $LicensesPerLoc.$LocCode) { #if there are no licenses assigned to this location code, we need to initialize
                        $HashToAdd = @{$License = 1} #set the count of this license to 1 as a hash value
                        $LicensesPerLoc.$LocCode = $HashToAdd #assign the hash into this hash. This is now a 2d hash
                    }
                    elseif (-not $CurrentCountOfLicenses) { #if the location does have some licenses, we don't need to initialize it, but just assign this particular license to count of 1.
                        $LicensesPerLoc.$LocCode.$License = 1
                    }
                    else {
                        $LicensesPerLoc.$LocCode.$License = $CurrentCountOfLicenses + 1 #if there is this specific license at this specific location code, we need to increment by 1.
                    }
                    #Write-Output "Unable to count office license for user $Upn. They either cannot be found or don't exist in AD.`nSpecifically, This means that No UserPrincipalName was ever found in AD Matching $Upn" | Out-File -Append -FilePath "C:\ErrorsInOfficeLicenses.txt"
                    #Write-Output "$Upn : $License (Not displayed in Unknown)" | Out-File -Append -FilePath "C:\ErrorsInOfficeLicenses.txt"
                    continue
                }
                if (-not $License) { #if not one of the defined licenses or.
                    continue
                }
                $LocCode = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData "LocationCode" #Get Location code for this user.
                if (-not $LocCode) { #if we don't have a valid location code
                    $OutputObj = @{
                        "User" = $Upn;
                        "License" = $License;
                        "Reason for Unknown" = "No Department Number Attribute"
                    }
                    $OutputObject = [PSCustomObject]$OutputObj
                    $OutputObject | Export-Csv -Path $SavePath -Append -NoTypeInformation -Force
                    $LocCode = "Unknown"
                    $CurrentCountOfLicenses = $LicensesPerLoc.$LocCode.$License #load current number of this license per this location
                    if (-not $LicensesPerLoc.$LocCode) { #if there are no licenses assigned to this location code, we need to initialize
                        $HashToAdd = @{$License = 1} #set the count of this license to 1 as a hash value
                        $LicensesPerLoc.$LocCode = $HashToAdd #assign the hash into this hash. This is now a 2d hash
                    }
                    elseif (-not $CurrentCountOfLicenses) { #if the location does have some licenses, we don't need to initialize it, but just assign this particular license to count of 1.
                        $LicensesPerLoc.$LocCode.$License = 1
                    }
                    else {
                        $LicensesPerLoc.$LocCode.$License = $CurrentCountOfLicenses + 1 #if there is this specific license at this specific location code, we need to increment by 1.
                    }
                    continue
                }
                $CurrentCountOfLicenses = $LicensesPerLoc.$LocCode.$License #load current number of this license per this location
                if (-not $LicensesPerLoc.$LocCode) { #if there are no licenses assigned to this location code, we need to initialize
                    $HashToAdd = @{$License = 1} #set the count of this license to 1 as a hash value
                    $LicensesPerLoc.$LocCode = $HashToAdd #assign the hash into this hash. This is now a 2d hash
                }
                elseif (-not $CurrentCountOfLicenses) { #if the location does have some licenses, we don't need to initialize it, but just assign this particular license to count of 1.
                    $LicensesPerLoc.$LocCode.$License = 1
                }
                else {
                    $LicensesPerLoc.$LocCode.$License = $CurrentCountOfLicenses + 1 #if there is this specific license at this specific location code, we need to increment by 1.
                }
            }
        }
    }
    #Write-Host "See the Error File C:\ErrorsInOfficeLicenses.txt to see if any errors in license + Location lookup were found.`n`nDATA BELOW`n==========" -ForegroundColor Yellow
    Write-Host "Finished Building Data Structure and writing unknown users to CSV." -ForegroundColor Green
    #All users without known location are now written to the CSV file.
    Return $LicensesPerLoc #2D Hash representing the number of licenses per location code.
}

#Takes in a 2D hash that represents the number of licenses per location code.
#Writes the data to Excel.
Function WriteLicenseDataToExcel ($HashOfLicenseData) {
    $FilePath = $env:USERPROFILE + "\Desktop\License&LocationData.xlsx"
    Write-Host "`nGenerating License&LocationData.xlsx saved to:`n" -ForegroundColor Magenta
    Write-Host "    $FilePath`n" -ForegroundColor Cyan
    #Stores column value of the license
    $LicenseExcelLocations = @{
        "Microsoft Stream Trial" = 2;
        "Project Plan 3" = 3;
        "Power BI Free" = 4;
        "Microsoft Teams Exploratory" = 5;
        "Flow Free" = 6;
        "F3" = 7;
        "Exchange Online Kiosk" = 8;
        "E3" = 9;
        "E1" = 10;
        "Visio Plan 2" = 11;
        "Power Automate per user with attended RPA Plan" = 12;
        "Microsoft Business Center" = 13;
        "Power BI Pro" = 14
    }
    if (Test-Path $FilePath) {
        Remove-Item $FilePath -Force
    }
    "" | Export-Excel $FilePath -WorksheetName "LicenseData" #Build empty excel file
    $RowCounter = 2
    $ExcelPkgFile = Open-ExcelPackage -Path $FilePath #Open Excel file
    $WorkSheet= $ExcelPkgFile.Workbook.Worksheets["LicenseData"] #open excel workbook
    $count = 2
    ForEach ($KVPair in $LicenseNames.GetEnumerator()) {
        $LicenseName = $KVPair.Value
        $WorkSheet.Cells[1,$count].Value = $LicenseName #write data to excel
        $count += 1
    }
    $WorkSheet.Cells[$RowCounter, 1].Value = $LocationCode #write data to excel
    #outer hash is the location code.
    ForEach ($Key in ($HashOfLicenseData.GetEnumerator())) {
        $WorkSheet.Cells[$RowCounter,1].Value = $Key.Name #Write the location code
        $InnerHash = $Key.Value #store in variable to make iterating more readable
        #nested key is the license and the count
        ForEach ($NestedKey in ($InnerHash.GetEnumerator())) {
            $ColumnVal = $LicenseExcelLocations."$($NestedKey.Name)" #Find the license column value. Using the hash hardcoded above.
            $WorkSheet.Cells[$RowCounter,$ColumnVal].Value = $NestedKey.Value #Write count of licenses to correct row and location.
        }
        $RowCounter += 1 #increment the row to write the new location code below previous.
    }
    Write-Host "Saving File..." -ForegroundColor Magenta
    Close-ExcelPackage $ExcelPkgFile #close and save the file when finished
    Write-Host "Completed!" -ForegroundColor Green
}
#Takes in nothing
#Acts as a frontEnd for the function GetOfficeLicenseDataPerLocation. It will output the data to the console.
Function RunOfficeDataReport {
    Clear-Host
    $OfficeData = GetOfficeLicenseDataPerLocation
    ForEach ($Key in ($OfficeData.GetEnumerator())) {
        $Location = $Key.Name
        $Hash = $Key.Value
        Write-Host -NoNewline -ForegroundColor Yellow "`nLocation Code: "
        Write-Host -NoNewline -ForegroundColor Cyan $Location
        $Dept = GetDeptNameForLocation $Location
        Write-Host " -- " -ForegroundColor Yellow -NoNewline
        Write-Host $Dept -ForegroundColor Cyan
        ForEach ($NestedKey in ($Hash.GetEnumerator())) {
            $LicenseName = $NestedKey.Name 
            $Count = $NestedKey.Value
            Write-Host -NoNewline -ForegroundColor Yellow "    License: "
            Write-Host -NoNewline -ForegroundColor Cyan $LicenseName
            Write-Host -NoNewline -ForegroundColor Yellow " : "
            Write-Host -ForegroundColor Cyan $Count
        }
    }
    cmd /c pause
}

#Takes in an email address (Upn)
#Returns a boolean on whether the user has an office license.
Function DoesUserHaveLicense($EmailAddress) {
    ConnectMSOL
    $flag = $false
    (Get-MsolUser -UserPrincipalName $EmailAddress -ErrorAction SilentlyContinue).licenses.AccountSkuId | ForEach-Object {
        if ($_) {
            $flag = $true
            if ($flag) { #just to get rid of dumb VSCODE warning of unused variable. 
                return $flag
            } 
        }
    }

    return $flag
}

#Takes in a location code
#Returns to the console, each user and their office license as well as the count of the totals for that location.
Function DisplayOfficeLicensesAtLocation ($LocationCode) {
    ConnectMSOL
    $LicenseData = @{ } #Hash we will store totals in.
    $LocationCode = ValidateLocationCode $LocationCode
    $ListOfUsers = GetAllUsersAtLocation $LocationCode
    Write-Host "`nOffice License Information at Location $LocationCode" -ForegroundColor Yellow
    Write-Host "==========================================`n" -ForegroundColor Yellow
    
    ForEach ($User in $ListOfUsers) { #loop through all users
        $Upn = GetADUserCustom -SearchBy sAMAccountName -SearchFor $User -ReturnData Mail
        if ($null -eq $Upn) {
            continue
        }
        (Get-MsolUser -UserPrincipalName $Upn).licenses.AccountSkuId | ForEach-Object { #Get all the user's office licenses
            if ($null -eq $_) { #if no office license, we don't care. Skip to next user.
                continue
            }
            $LicName = $LicenseNames.$_
            Write-Host "User: " -NoNewline -ForegroundColor Yellow
            Write-Host "$User" -ForegroundColor Cyan -NoNewline
            Write-Host " has license " -ForegroundColor Yellow -NoNewline
            Write-Host $LicName -ForegroundColor Cyan
            $CurrentCountOfLicenses = $LicenseData.$LicName #load current number of this license per this location
            if (-not $CurrentCountOfLicenses) { #if the location does have some licenses, we don't need to initialize it, but just assign this particular license to count of 1.
                $LicenseData.$LicName = 1 #initialize license count
            }
            else {
                $LicenseData.$LicName = $CurrentCountOfLicenses + 1 #increment license count.
            }
        }
    }
    Write-Host "`nTotals" -ForegroundColor Yellow
    Write-Host "======" -ForegroundColor Yellow
    ForEach ($Key in ($LicenseData.GetEnumerator())) {
        $LicenseName = $Key.Name
        $LicenseCount = $Key.Value
        Write-Host "$LicenseName : $LicenseCount" -ForegroundColor Cyan
    }
}

#START OF MAILBOX MANAGEMENT FUNCTIONS

#Takes in an employee ID
#Takes in a mailbox name
#Assigns the employee to have the mailbox permission(full access)
Function AssignMailboxPermissions($EmployeeID, $Mailbox) {
    ConnectEXO
    $err1 = $null
    $EmpName = GetADUserCustom -SearchBy EmployeeID -SearchFor $EmployeeID -ReturnData Name
    $MailboxID = Get-Mailbox -Anr $Mailbox -ErrorAction SilentlyContinue | Select-Object ExternalDirectoryObjectId
    if ($null -eq $MailboxID) {
        Write-Host "Unable to find Mailbox `"$Mailbox`" Make sure it is spelled correctly in Access Matrix." -ForegroundColor Red
        return
    }
    else {
        $Counter = 0
        $UserReadyInEXO = IsUserShownInEXO $EmpName
        while ($UserReadyInEXO -eq $false) {
            if ($Counter -eq 0) {
                Write-Host "Waiting for user to be shown in Exchange Online.`nThis may take a while for new users." -ForegroundColor Magenta -NoNewline
            }
            $UserReadyInEXO = IsUserShownInEXO $EmpName
            $Counter = $Counter + 1
            if ($Counter % 20 -eq 0) { #output a . every 20 tries to show the user it's still working.
                Write-Host -NoNewline "." -ForegroundColor Magenta
            }
        }
        if ($Counter -gt 0) {
            Write-Host ""
        }
        Add-MailboxPermission -Identity $MailboxID."ExternalDirectoryObjectId" -User $EmpName -AccessRights FullAccess -InheritanceType All -ErrorVariable $err1 | Out-Null
    } 
    if ($null -ne $err1) {
        return -1
    }
    Write-Host "Added " -ForegroundColor Green -NoNewline
    Write-Host "$EmpName" -ForegroundColor Cyan -NoNewline
    Write-Host " to have permissions to mailbox " -NoNewline -ForegroundColor Green
    Write-Host "$Mailbox" -NoNewline -ForegroundColor Cyan
    Write-Host " successfully." -ForegroundColor Green
}

#Takes in an employeeID
#Removes the users permission to all mailboxes.
Function RemoveMailboxPermissions($EmployeeID) {
    ConnectEXO
    $err1 = $null
    $EmpName = GetADUserCustom -SearchBy EmployeeID -SearchFor $EmployeeID -ReturnData Name
    $sAM = GetADUserCustom -SearchBy Name -SearchFor $EmpName -ReturnData "sAMAccountName"
    $PossibleMailboxMembership = GetAllDeptMailboxes #Retreives the Dept Mailboxes from the OU in AD.
    $MailboxNames = [System.Collections.ArrayList]@()
    ForEach ($MailboxSAM in $PossibleMailboxMembership) { #this loop retreives all the licensed mailboxes in our AD environment.
        $EmailAddress = GetADUserCustom -SearchBy sAMAccountName -SearchFor $MailboxSAM -ReturnData Mail
        if ($null -eq $EmailAddress) {
            continue
        }
        $HasLicense = DoesUserHaveLicense $EmailAddress
        if ($HasLicense) {
            $MailboxNames.Add($MailboxSAM) | Out-Null
        }
    }
    $SharedMailboxes = GetSharedMailboxPermissions $sAM
    ForEach ($Mailbox in $SharedMailboxes) {
        $MailboxNames.Add($Mailbox."Identity") | Out-Null
    }
    ForEach ($MailboxIter in $MailboxNames) { #loop through mailboxes and find if a user has permissions to any.
        $Result = Get-Mailbox -Anr $MailboxIter | Get-MailboxPermission -User $EmpName
        if ($Result) { #if result is not null, the user has permissions
            $MailboxID = Get-Mailbox -Anr $MailboxIter | Select-Object ExternalDirectoryObjectId
            $MailboxID = $MailboxID."ExternalDirectoryObjectId"
            Remove-MailboxPermission -Identity $MailboxID -User $EmpName -AccessRights FullAccess -ErrorVariable $err1 -Confirm:$false
            if ($null -ne $err1) {
                return -1
            }   
            else {
                Write-Host "Successfully removed " -ForegroundColor Green -NoNewline
                Write-Host "$EmpName" -ForegroundColor Cyan -NoNewline
                Write-Host " from mailbox: " -NoNewline -ForegroundColor Green
                Write-Host "$MailboxIter" -ForegroundColor Cyan
            }
        }
    } 
}

#Takes in a SAMACcountNAme
#Returns the list of permissions to mailboxes.
Function GetSharedMailboxPermissions($sAMAccountName) {
    ConnectEXO
    $EmailAddr = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $MailboxMembership = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission | Select-Object Identity,User  | Where-Object {($_.user -like $EmailAddr)}
    return $MailboxMembership
}


#Taken from Microsoft Docs 
#Sets the MFA requirement state
Function Set-MfaState {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        $ObjectId,
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        $UserPrincipalName,
        [ValidateSet("Disabled","Enabled","Enforced")]
        $State
    )
    Process {
        Write-Verbose ("Setting MFA state for user '{0}' to '{1}'." -f $ObjectId, $State)
        $Requirements = @()
        if ($State -ne "Disabled") {
            $Requirement =
                [Microsoft.Online.Administration.StrongAuthenticationRequirement]::new()
            $Requirement.RelyingParty = "*"
            $Requirement.State = $State
            $Requirements += $Requirement
        }
        Set-MsolUser -ObjectId $ObjectId -UserPrincipalName $UserPrincipalName `
                     -StrongAuthenticationRequirements $Requirements
    }
}

Function DisableMFA ($sAMAccountName) {
    ConnectMSOL
    $UserPrincipalName = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
    $ObjID = $User."ObjectID"
    Set-MfaState -ObjectId $ObjID -UserPrincipalName $UserPrincipalName -State "Disabled"
}

Function EnableMFA ($sAMAccountName) {
    ConnectMSOL
    $UserPrincipalName = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
    $ObjID = $User."ObjectID"
    Set-MfaState -ObjectId $ObjID -UserPrincipalName $UserPrincipalName -State "Enabled"
}

Function EnforceMFA ($sAMAccountName) {
    ConnectMSOL
    $UserPrincipalName = GetADUserCustom -SearchBy sAMAccountName -SearchFor $sAMAccountName -ReturnData Mail
    $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
    $ObjID = $User."ObjectID"
    Set-MfaState -ObjectId $ObjID -UserPrincipalName $UserPrincipalName -State "Enforced"
}