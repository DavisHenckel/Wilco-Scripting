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