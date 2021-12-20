Function GetADUserCustom {
    <#
    .SYNOPSIS
        Gets a specific portion of a commonly used Active Directory Attribute.
    .DESCRIPTION
        Issues a Get-ADUser command in a format that is a bit more intuitive and narrows down the scope to just commonly used attributes for Wilco.
    .PARAMETER SearchBy
        The attribute to search for the user by. This must be in the set "EmployeeID", "Name", "sAMAccountName", "Mail", "DistinguishedName" 
    .PARAMETER SearchFor
        The value to lookup the AD User by 
    .PARAMETER ReturnData
        The attribute(s) to return. These are commonly used attributes for Wilco. All Does not return all of the AD attributes. It represents the AD Attributes commonly used at Wilco. ReturnData must be in the set "EmployeeID", "Name", "Department", "sAMAccountName", "Mail", "LocationCode", "DistinguishedName", "Title", "Mobile", "Manager", "All". If no value is specified, "All" is used.
    .EXAMPLE
        $LogonName = GetADUserCustom -SearchBy Name -SearchFor "Davis Henckel" -ReturnData sAMAccountName
    .EXAMPLE
        $UserInfo = GetADUserCustom -SearchBy sAMAccountName -SearchFor "Davis Henckel"
    .OUTPUTS
        Returns either the specific value specified in the ReturnData Attribute, or, if ReturnData is "All", returns a HashTable containing all commonly used attributes at Wilco.
    #>
    param(
        [parameter(Mandatory=$true)]
        [ValidateSet("EmployeeID", "Name", "sAMAccountName", "Mail", "DistinguishedName")]
        [string] $SearchBy,
        [parameter(Mandatory=$true)]
        [string] $SearchFor,
        [parameter(Mandatory=$false)]
        [ValidateSet("EmployeeID", "Name", "Department", "sAMAccountName", "Mail", "LocationCode", "DistinguishedName", "Title", "Mobile", "Manager", "All")]
        [string] $ReturnData = "All"
    )
    if ($ReturnData -eq "LocationCode") {
        $ReturnData = "departmentNumber"
    }
    $User = $null
    #Capture the User.
    if ($SearchBy -eq "EmployeeID") {
        $User = Get-ADUser -Filter {employeeID -eq  $SearchFor -and Enabled -eq $true} -Properties *
    }
    elseif ($SearchBy -eq "Name") {
        $User = Get-ADUser -Filter {Name -eq  $SearchFor -and Enabled -eq $true} -Properties *
    }
    elseif ($SearchBy -eq "sAMAccountName") {
        $User = Get-ADUser -Filter {sAMAccountName -eq  $SearchFor -and Enabled -eq $true} -Properties *
    }
    elseif ($SearchBy -eq "Mail") {
        $User = Get-ADUser -Filter {mail -eq  $SearchFor -and Enabled -eq $true} -Properties *
    }
    elseif ($SearchBy -eq "DistinguishedName") {
        $User = Get-ADUser -Filter {DistinguishedName -eq  $SearchFor -and Enabled -eq $true} -Properties *
    }
    if ($null -eq $User) {
        Write-Host "User was not found. Searched for: $SearchFor and searched by $SearchBy" -ForegroundColor Red
        return
    }
    if ($ReturnData -eq "All" -or $ReturnData -eq "all") {
        $DataToReturn = @{
            "Location" = $User.departmentNumber[0];
            "Department" =  $User.Department
            "Employee Number" = $User.EmployeeID;
            "Name" = $User.Name;
            "Email Address" = $User.mail;
            "Logon Name" = $User.SamAccountName;
            "Job Title" = $User.Title;
            "Distinguished Name" = $User.DistinguishedName;
            "Mobile Number" = $User.Mobile
            "Manager" = GetNameFromDN($User.Manager)
        }
        Return $DataToReturn
    }
    else {
        Return $User.$ReturnData
    }
}

Function GetLocationForDept {
    <#
    .SYNOPSIS
        Gets a 3 digit location code.
    .DESCRIPTION
        GetLocationForDept uses the location reference list "LocationInfo" worksheet to return a string containing the 3 digit location code for a department name. 
    .PARAMETER DeptName
        Contains a department name IE: "Human Resources", "Marketing", etc.
    .EXAMPLE
        $LocationCode = GetLocationForDept "Marketing"
    .OUTPUTS
        Returns a 3 digit location code in string format. Output is given in string format to prevent issues with location codes starting in 0.
    #>
    param(
        [parameter(Mandatory=$true)]
        [String]$DeptName
    )
    $PathToFile = $PSScriptRoot + "\Excel Dependencies\LocationFile.xlsx"
    
    $ExcelFile = $null
    Try {
        $ExcelFile = Import-Excel $PathToFile -WorksheetName "LocationInfo"
    }
    Catch {
        Write-Host "ERROR -- Could not read file at path $($PathToFile). Exiting" -ForegroundColor Red
        cmd /c pause
        exit
    }
    $ArrayOfDepts = $ExcelFile."Department"
    $IndexOfDept = $ArrayOfDepts.IndexOf($DeptName)
    $DataForLocation = $ExcelFile[$IndexOfDept]
    if ($IndexOfDept -eq -1) {
        Write-Host "Unable to find $DeptName" -ForegroundColor Red
        return $null
    }
    $Ret = $DataForLocation."departmentNumber"
    return $Ret    
}

Function GetLocationInfo {
    <#
    .SYNOPSIS
        Gets data for a location, returned in a hash table.
    .DESCRIPTION
        GetLocationInfo uses the location reference list "LocationInfo" worksheet to return the row of all information that corresponds to the 3 digit location code for a department name. 
    .PARAMETER LocationCode
        Contains a department number IE: "010", "020", etc.
    .EXAMPLE
        $LocationData = GetLocationInfo "010"
    .OUTPUTS
        Returns a hashtable that contains the row of the corresponding location code. This can be used to assign various AD Attributes.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$LocationCode
    )
    $LocationCode = "$LocationCode"
    if ($LocationCode.Length -eq 2) {
        $LocationCode = "0" + $LocationCode
    }
    $PathToFile = $PSScriptRoot + "\Excel Dependencies\LocationData.xlsx"
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
    $DataForLocation = $ExcelFile[$IndexOfLocationCode]
    $DataHash = @{} #This is a Hash table we will enumerate
    $DataForLocation.psobject.properties | ForEach-Object { $DataHash[$_.Name] = $_.Value } #map the custom PS Object to a PS hash table.
    if ($IndexOfLocationCode -eq -1) {
        Write-Host "Unable to find Location Code `"$LocationCode`"" -ForegroundColor Red
    }
    Return $DataHash #Returns a hash table containing the row of the excel file.
}

Function GetOUForLocation {
    <#
    .SYNOPSIS
        Gets the OU for a given location code.
    .DESCRIPTION
        Returns the OU for a given location code. See GetLocationInfo for specifics. It uses the LocationInfo worksheet and returns the OU.
    .PARAMETER LocNum
        Contains a department number IE: "010", "020".
    .EXAMPLE
        $OU = GetOUForLocation "010"
    .OUTPUTS
        Returns a string that contains the OU for a given location.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$LocNum
    )
    $LocNum = ValidateLocationCode($LocNum) #reassigns $LocNum to a valid Location number. Ideally it would be validated before this.
    $LocInfo = GetLocationInfo $LocNum
    $LocationOU = $LocInfo."OU"
    return $LocationOU #return it
}

Function GetDeptNameForLocation {
    <#
    .SYNOPSIS
        Gets the Department Name for a given location code.
    .DESCRIPTION
        Returns the Department Name for a given location code. See GetLocationInfo for specifics. It uses the LocationInfo worksheet and returns the department name.
    .PARAMETER LocNum
        Contains a department number IE: "010", "020".
    .EXAMPLE
        $DeptName = GetDeptNameForLocation "010"
    .OUTPUTS
        Returns a string that contains the department name for a given location.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$LocNum
    )
    $LocNum = ValidateLocationCode($LocNum) #reassigns $LocNum to a valid Location number. Ideally it would be validated before this.
    $LocInfo = GetLocationInfo $LocNum
    $LocationCity = $LocInfo."Department"
    return $LocationCity #return it
}

#Takes in nothing
#Returns an arraylist of mailbox names that are defined in the dept mailboxes OU
Function GetAllDeptMailboxes {
    <#
    .SYNOPSIS
        Gets a subset of online mailboxe within a certain OU.
    .DESCRIPTION
        Searches an OU for AD Email names, then searches in AzureAD to see if they are licensed and in use.
    .EXAMPLE
        $Mailboxes = GetAllDeptMailboxes
    .OUTPUTS
        Returns an ArrayList that contains the sAMAccountNames of each mailbox.
    #>
    $Mailboxes = Get-AdUser -Filter * -SearchBase "SpecificOUPath" | Select-Object sAMAccountName
    $MailboxNames = [System.Collections.ArrayList]@()
    ForEach ($Mailbox in $Mailboxes) {
        $MailboxNames.Add($Mailbox."sAMAccountName") | Out-Null
    }
    return $MailboxNames
}

Function GetCityNameForLocation {
    <#
    .SYNOPSIS
        Gets the city name for a given location code.
    .DESCRIPTION
        Returns the city name for a given location code. See GetLocationInfo for specifics. It uses the LocationInfo worksheet and returns the city name.
    .PARAMETER LocNum
        Contains a department number IE: "010", "020".
    .EXAMPLE
        $CityName = GetDeptNameForLocation "010"
    .OUTPUTS
        Returns a string that contains the department name for a given location.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$LocNum
    )
    $LocNum = ValidateLocationCode($LocNum) #reassigns $LocNum to a valid Location number. Ideally it would be validated before this.
    $LocInfo = GetLocationInfo $LocNum
    $LocationCity = $LocInfo."City"
    return $LocationCity #return it
}


Function GetAccessMatrixData {
    <#
    .SYNOPSIS
        Enumerates the spreadsheet that defines job titles.
    .DESCRIPTION
        Uses ImportExcel module to enumerate a PSCustomObject that contains the data of the worksheet that defines access levels for job titles.
    .PARAMETER PathToFile
        String that represents the path to the Excel file that contains the data to enumerate.
    .EXAMPLE
        $File = GetAccessMatrixData "C:\test.xlsx"
    .OUTPUTS
        Returns PSCustomObject that contains the job role data.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$PathToFile
    )
    Try {
        $AccessMatrixFile = Import-Excel $PathToFile -WorksheetName WorksheetName #Load the Worksheet
    }
    Catch {
        Write-Host "ERROR -- Could not read file at path $($PathToFile). Exiting" -ForegroundColor Red
        cmd /c pause
        exit
    }
    return $AccessMatrixFile
}

Function GetOfficeLicenseForJob  {
    <#
    .SYNOPSIS
        Gets an office license name based on a job title
    .DESCRIPTION
        Uses ImportExcel to read in the Access spreadsheet. Looks up the row of the job title, then retrieves the corresponding value in the office license field.
    .PARAMETER AccessMatrixFile
        PSCustomObject that contanis the data of the AccessMatrix.
    .PARAMETER JobTitle
        String that represents a job title.
    .EXAMPLE
        $License = GetOfficeLicenseForJob $AccessFile "Service Desk Technician"
    .OUTPUTS
        Returns a string that contains the office license data (Exchange online Kiosk, F3, E3, etc).
    #>
    param (
        [parameter(Mandatory=$true)]
        [PSCustomObject]$AccessMatrixFile,
        [parameter(Mandatory=$true)]
        [string]$JobTitle
    )
    $JobTitleArray = $AccessMatrixFile."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    if ($IndexOfJobTitle -eq -1) {
        Write-Host "WARNING -- This Job title is not defined in the access matrix. Would you like to re-enter the job?(y/n)`nYou can select no, but then You cannot get the AD container, or the OU to be placed in.`nEnter your selection (y/n): " -ForegroundColor Red -NoNewline
        while ($true) {
            $UserChoice = Read-Host
            if ($UserChoice -eq 'y' -or $UserChoice -eq 'Y') {
                $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
                $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the validated job title.
                break
            }
            elseif ($UserChoice -eq 'n' -or $UserChoice -eq 'N') {
                return $null
            }
            Write-Host "Invalid Input. Please enter `"y`" if you want to validate Job Title, `"n`" if you do not."
        }
        $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
    }
    if ($IndexOfJobTitle -eq -1) { #if the user chose to override an invalid job title yet again... come on man....
        Write-Host "Unable to validate Job title..." -ForegroundColor Red
        return $null
    }
    $DetailsOfJobTitle = $AccessMatrixFile[$IndexOfJobTitle] #This is a CustomPS Object, it is setup like a hash table but doesn't have any of the methods of a hash table.
    $LicenseName = $DetailsOfJobTitle."Office 365 License"
    if ($LicenseName -match "F3") {
        $LicenseName = "F3" #remove the (Formerly F1)
    }
    Return $LicenseName
}

Function GetADContainerForJob {
    <#
    .SYNOPSIS
        Gets an AD container description name based on a job title
    .DESCRIPTION
        Uses ImportExcel to read in the Access spreadsheet. Looks up the row of the job title, then retrieves the corresponding value in the AD Container field.
    .PARAMETER AccessMatrixFile
        PSCustomObject that contanis the data of the AccessMatrix.
    .PARAMETER JobTitle
        String that represents a job title.
    .EXAMPLE
        $License = GetADContainerForJob $AccessFile "Service Desk Technician"
    .OUTPUTS
        Returns a string that contains the AD Container for the given job.
    #>
    param (
        [parameter(Mandatory=$true)]
        [PSCustomObject]$AccessMatrixFile,
        [parameter(Mandatory=$true)]
        [string]$JobTitle
    )
    $JobTitleArray = $AccessMatrixFile."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    if ($IndexOfJobTitle -eq -1) {
        Write-Host "WARNING -- This Job title is not defined in the access matrix. Would you like to re-enter the job?(y/n)`nYou can select no, but then You cannot get the AD container, or the OU to be placed in.`nEnter your selection (y/n): " -ForegroundColor Red -NoNewline
        while ($true) {
            $UserChoice = Read-Host
            if ($UserChoice -eq 'y' -or $UserChoice -eq 'Y') {
                $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
                $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the validated job title.
                break
            }
            elseif ($UserChoice -eq 'n' -or $UserChoice -eq 'N') {
                return $null
            }
            Write-Host "Invalid Input. Please enter `"y`" if you want to validate Job Title, `"n`" if you do not."
        }
        $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
    }
    if ($IndexOfJobTitle -eq -1) { #if the user chose to override an invalid job title yet again... come on man....
        Write-Host "Unable to validate Job title..." -ForegroundColor Red
        return $null
    }
    $DetailsOfJobTitle = $AccessMatrixFile[$IndexOfJobTitle] #This is a CustomPS Object, it is setup like a hash table but doesn't have any of the methods of a hash table.
    $ADContainer = $DetailsOfJobTitle."AD Container"
    Return $ADContainer
}

Function GetOUForADContainer {
    <#
    .SYNOPSIS
        Gets an OU for a given AD Container.
    .DESCRIPTION
        Checks to see if the user is supposed to be an OWA User. Otherwise, gets the OU for the specific location if not an OWA user.
    .PARAMETER ADContainer
        A string that describes an OU.
    .PARAMETER LocationCode
        A 3 digit string that represents a location code. 
    .EXAMPLE
        $OU = GetOUForADContainer "Retail:Users:Location" "010"
    .OUTPUTS
        Returns an OU that corresponds to a valid OU in Wilco's Active Directory Structure.
    #>
    param (
        [parameter(Mandatory=$true)]
        [string]$ADContainer,
        [parameter(Mandatory=$true)]
        [string]$LocationCode
    )
    if ($null -eq $ADContainer) { 
        return "SpecificOUPath"
    }
    if ($ADContainer -match "OWA") {
        return "SpecificOUPath"
    }
    else {
        $OU = GetOUForLocation $LocationCode
        return $OU
    }
}

#Takes in an arraylist containing group names
#Returns another arraylist that have prefix "MB: " which indicates this is a mailbox.
Function GetMailboxesFromGroupList ($Groups) {
    $Mailboxes = [System.Collections.ArrayList]@()
    ForEach ($Group in $Groups) {
        if ($Group -match "MB: ") {
            $Group = $Group.SubString(4, $Group.Length - 4) #strip first 4 characters off.
            $Mailboxes.Add($Group) | Out-Null
        }
    }
    return $Mailboxes
}

# Takes in the Access Matrix File. This can be done easily by Calling GetAccessMatrixData and passing it the path to the Access Matrix
# Takes in a Job Title that should be defined in the Access Matrix
# Takes in a Location code.
# Returns an arraylist of Groups that are needed for The job. This is based of the data we see in the Access Matrix.
Function GetGroupsForJobTitle ($AccessMatrixFile, $JobTitle, $LocationCode) {
    $JobTitleArray = $AccessMatrixFile."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    if ($IndexOfJobTitle -eq -1) {
        Write-Host "WARNING -- This Job title is not defined in the access matrix. Would you like to re-enter the job?(y/n)`nYou can select no, but then Job Groups will not be setup.`nEnter your selection (y/n): " -ForegroundColor Red -NoNewline
        while ($true) {
            $UserChoice = Read-Host
            if ($UserChoice -eq 'y' -or $UserChoice -eq 'Y') {
                $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
                $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
                break
            }
            elseif ($UserChoice -eq 'n' -or $UserChoice -eq 'N') {
                return $null
            }
            Write-Host "Invalid Input. Please enter `"y`" if you want to validate Job Title, `"n`" if you do not."
        }
        $JobTitle = ValidateJobTitle($JobTitle) #validate the job title.
    }
    if ($IndexOfJobTitle -eq -1) {
        Write-Host "Unable to validate Job title..." -ForegroundColor Red
        return
    }
    $DetailsOfJobTitle = $AccessMatrixFile[$IndexOfJobTitle] #This is a CustomPS Object, it is setup like a hash table but doesn't have any of the methods of a hash table.
    $DetailsOfJobHashTable = @{} #This is a Hash table we will enumerate
    $DetailsOfJobTitle.psobject.properties | ForEach-Object { $DetailsOfJobHashTable[$_.Name] = $_.Value } #map the custom PS Object to a PS hash table.
    $City = GetCityNameForLocation($LocationCode) #get the city name based on location, so we can assign the correct retail DGs.
    $GroupsToAdd = [System.Collections.ArrayList]@() #create arraylist to contain groups the users need based on data from the hash table. This is what we return at the end of function
    ForEach ($Key in ($DetailsOfJobHashTable.GetEnumerator() | Where-Object {$_.Value -eq "X" -or $_.Value -match "Location" -or $_.Value -eq "Setup: Email Controller" -or $_.Value -eq "All Locations" -and ($_.Name.Contains('MB:') -or $_.Name.Contains('SG:') -or $_.Name.Contains('DG:'))})) {
        $NameOfGroup = $Key.Name
        if ($Key.Name -match "MB:" -eq 0) { #if it's a mailbox. Keep Identifier to know to try assigning in O365 later on.
            $NameOfGroup = $Key.Name.SubString(4, $Key.Name.Length - 4) #otherwise take off the first 4 chars ("DG: ") so we can attempt group assignment later on.
        }
        if ($NameOfGroup -match "Retail Location Group") {
            $TempCity = $City.Replace(" ", "")
            $NameOfGroup = $NameOfGroup.Replace("(Retail Location Group)", $TempCity)
        }
        if ($NameOfGroup -match "Site Distribution Group" -and $Key.Value -ne "All Locations") {
            $NameOfGroup = "Retail $($City)"
        }
        if ($NameOfGroup -eq "Domain Users") {
            continue
        }
        $GroupsToAdd.Add($NameOfGroup) | Out-Null
    }
    return $GroupsToAdd
}

#Takes in a distinguishedName
#Returns a string containing the users OU.
Function GetUsersCurrentOU ($UsersDistName) {
    $UsersDistName = $UsersDistName.Substring($UsersDistName.IndexOf(',') + 1,$UsersDistName.Length - $UsersDistName.IndexOf(',') - 1) #Slice off CN=... so we are left with only their OU
    return $UsersDistName
}

#Takes in nothing
#Returns a hashtable of users that are hid from the GAL that exist in the retail OU.
Function GetUsersHidFromGALRetailOU {
    $Users = Get-ADUser -Filter {msExchHideFromAddressLists -eq $true} -SearchBase 'SpecificOUPath' | Select-Object Name, DistinguishedName
    #Write-Host $Users
    return $Users
}

#Takes in a job title
#Returns an arraylist of all users with that job title.
Function GetAllUsersWithJobTitle ($JobTitle) {
    $ReturnArrayList = [System.Collections.ArrayList]@() 
    $Users = Get-ADUser -Filter {title -eq $JobTitle -and Enabled -eq $true} -Properties sAMAccountName, name | Select-Object sAMAccountName
     ForEach ($User in $Users) {
        $ReturnArrayList.Add($User."sAMAccountName") | Out-Null
    }
    return $ReturnArrayList
}

#Takes in a distinguishedname
#Returns a list of users that have the manager of the distinguishedname
Function GetAllUsersWithManager ($ManagerDistinguishedName) {
    $ReturnArrayList = [System.Collections.ArrayList]@() 
    Try {
        $Users = Get-ADUser -Filter {manager -eq $ManagerDistinguishedName -and Enabled -eq $true} -Properties sAMAccountName | Select-Object sAMAccountName
        ForEach ($User in $Users) {
            $ReturnArrayList.Add($User."sAMAccountName") | Out-Null
        }
    }
    Catch {
        Write-Host "ERROR -- Unable to get employees with `"$ManagerDistinguishedName`" as their manager" -ForegroundColor Red
    }
    return $ReturnArrayList
}

#Takes in a SAMAccountName
#Returns an arraylist containing groups that the user has in AD.
Function GetUsersGroups ($sAMAccountName) {
    Try {
        $sAMAccountName = $sAMAccountName.Replace(" ", ".")
        $UserGroups = (Get-ADUser $sAMAccountName -Properties MemberOf).MemberOf
    }
    Catch {
        Write-Host "Error -- Unable to locate `"$sAMAccountName`"" -ForegroundColor Red
    }
    $ReadableGroups = [System.Collections.ArrayList]@()  
    ForEach ($Group in $UserGroups) {
        $Group = $Group.SubString(0, $Group.IndexOf(','))
        $Group = $Group.SubString(3, $Group.Length - 3)
        $ReadableGroups.Add($Group) | Out-Null
        #Write-Host $Group
    }
    $ReadableGroups.Add("Domain Users") | Out-Null
    return $ReadableGroups
}

#Takes in a location code.
#Returns an arraylist of users' SAMAccountNames at that location.
Function GetAllUsersAtLocation ($LocationCode) {
    $LocationCode = "$LocationCode"
    $ReturnArrayList = [System.Collections.ArrayList]@() 
    $Users = Get-ADUser -Filter {departmentNumber -eq $LocationCode -and Enabled -eq $true} -Properties sAMAccountName | Select-Object sAMAccountName
    ForEach ($User in $Users) {
        $ReturnArrayList.Add($User."sAMAccountName") | Out-Null
    }
    Return $ReturnArrayList
}

#Takes in a distinguishedname
#Returns a string that is the name of the user with that distinguishedname
Function GetNameFromDN ($DN) {
    try {
        $User = Get-ADUser -Filter {DistinguishedName -eq  $DN -and Enabled -eq $true} -Property name | Select-Object name
    }
    catch {
        Write-Host "Unable to locate `"$DN`" in AD." -ForegroundColor Red
        return
    }
    return $User."name"
}

#Takes in an employeeID
#Returns the users managers name.
Function GetUsersManager ($EmployeeID) {
    Try {
        $User = Get-ADUser -Filter {employeeID -eq $EmployeeID -and enabled -eq $true} -Property manager | Select-Object manager
    }
    Catch {
        Write-Host "ERROR -- Unable to locate `"$EmployeeID`" in AD." -ForegroundColor Red
    }
    $Name = GetNameFromDN $User."manager"
    return $Name
}

#Takes in a SAMAccountName
#Returns user password information to the console.
Function GetPasswordExpiration ($sAMAccountName) {
    Try {
        $ReturnValue = Net User $sAMAccountName /domain
    }
    Catch {
        Write-Host "ERROR -- Could not find `"$UserName`" in AD" -ForegroundColor Red
        return $null 
    }
    Write-Host -ForegroundColor Cyan "$($ReturnValue[3])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[2])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[10])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[11])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[12])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[20])"
    Write-Host -ForegroundColor Cyan "$($ReturnValue[14])"
}

#Takes in a job title [string]
#Takes in a location code [int or string]
#Returns a specific user at the location with the given job title.
Function GetUserAtLocationWithJobTitle ($JobTitle, $LocationCode) {
    $UsersAtThisLocation = GetAllUsersAtLocation $LocationCode
    $UsersWithJobTitle = [System.Collections.ArrayList]@()
    if ($JobTitle -match "Manager" -eq 0) {
        $JobTitle = ValidateJobTitle($JobTitle) #ensure we are searching for a valid job title
    }
    ForEach ($User in $UsersAtThisLocation) {
        Try {
            $UserAccount = Get-ADUser -Filter {sAMAccountName -eq  $User -and Enabled -eq $true} -Properties title | Select-Object title
        }
        Catch {
            Write-Host "Unable to locate `"$User`" in AD." -ForegroundColor Red
            continue
        }
        if ($UserAccount."title" -match $JobTitle) {
            $UsersWithJobTitle.Add($User) | Out-Null
        }
    }
    if ($UsersWithJobTitle.Count -gt 1) {
        Write-Host -ForegroundColor Yellow "WARNING -- There are more than 1 users with $JobTitle as their job title at $LocationCode."
        Write-Host -ForegroundColor Cyan "Which one would you like to return?"
        while($true) {
            $Counter = 1
            ForEach ($SpecificUser in $UsersWithJobTitle) {
                Write-Host "$Counter : $SpecificUser" -ForegroundColor Cyan
                $Counter = $Counter + 1
            }
            try { [int32]$UsersInput = Read-Host }
            catch {
                Write-Host "Must enter a number..." -ForegroundColor Red
                Start-Sleep 2
                continue #prompt again
            }
            if ($UsersInput -lt 1 -or $UsersInput -gt $Counter - 1) {
                Write-Host "Must enter number between 1 and $($Counter - 1)  (inclusive)"
                continue
            }
            else {
                return $UsersWithJobTitle[$UsersInput - 1]
            }
        }
    }
    else {
        return $UsersWithJobTitle[0]
    }
}

#Takes in a number of days
#Returns a hashtable to the console that shows the AD Users created within the last n number of days.
Function GetADUsersCreatedInPastDays ($NumDays) {
    $DateCutOff=(Get-Date).AddDays(-$NumDays)
    Try {
        Get-ADUser -Filter * -Property whenCreated | Where-Object {$_.enabled -eq $true -and $_.whenCreated -gt $datecutoff} | Format-Table Name, whenCreated -Autosize
    }
    Catch {
        Write-Host "Unable to run the command. Be sure to pass this function a day. IE to see past 10 days of created users, enter 10." -ForegroundColor Red
    }
}

#Takes in an access matrix custom PS Object
#Takes in a job title
#Returns a boolean that shows whether the job is a manager.
Function IsJobAManager ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    if ($IndexOfJobTitle -eq -1) {
        #Write-Host "$JobTitle is not a valid job defined in the access matrix. Unable to tell if this job is defined as a manager." -ForegroundColor Red
    }
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] #This is a CustomPS Object, it is setup like a hash table but doesn't have any of the methods of a hash table.
    $ManagerChange = $DetailsOfJobTitle."Manager Change"
    if ($ManagerChange) {
        return $true
    }
    return $false
}

#Takes in a location code
#Returns a manager name at the corresponding location code.
Function GetManagerAtLocation ($LocationCode) {
    $Managers = [System.Collections.ArrayList]@() 
    $LocationData = GetLocationInfo $LocationCode #Loads the row from Location Reference List
    $LocationMgrName = $LocationData."Manager"
    $LocationMgrSAM = GetADUserCustom -SearchBy Name -SearchFor $LocationMgrName -ReturnData "sAMAccountName"
    $AMPath = $PSScriptRoot + "\Excel Dependencies\AccessMatrix.xlsm"
    $AM = GetAccessMatrixData $AMPath
    $ArrayOfUsers = GetAllUsersAtLocation $LocationCode
    ForEach ($User in $ArrayOfUsers) {
        $UserJobTitle = GetADUserCustom -SearchBy sAMAccountName -SearchFor $User -ReturnData "Title"
        if (IsJobAManager $AM $UserJobTitle) {
            $Managers.Add($User) | Out-Null
        }
    }
    if ($Managers.Contains($LocationMgrSAM) -eq $false) {
        $Managers.Add($LocationMgrSAM) | Out-Null
    }
    if ($Managers.Count -gt 1) {
        Write-Host "There are more than 1 user at this location considered to be a manager. Which one would you like to assign?" -ForegroundColor Yellow
        while($true) {
            $Counter = 1
            ForEach ($SpecificUser in $Managers) {
                Write-Host "$Counter : $SpecificUser" -ForegroundColor Cyan
                $Counter = $Counter + 1
            }
            try { [int32]$UsersInput = Read-Host }
            catch {
                Write-Host "Must enter a number..." -ForegroundColor Red
                Start-Sleep 2
                continue #prompt again
            }
            if ($UsersInput -lt 1 -or $UsersInput -gt $Counter - 1) {
                Write-Host "Must enter number between 1 and $($Counter - 1)  (inclusive)"
                continue
            }
            else {
                return $Managers[$UsersInput - 1]
            }
        }
    }
    else {
        return $Managers[0]
    }
}
