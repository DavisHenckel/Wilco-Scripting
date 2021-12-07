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