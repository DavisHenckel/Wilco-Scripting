#Returns a Boolean
Function DoesJobGetEagle ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $Eagle = $DetailsOfJobTitle."Eagle Role"
    if ($Eagle) {
        return $true
    }
    return $false

}

#Returns a Boolean
Function DoesJobGetMOL ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $MOL = $DetailsOfJobTitle."MOL Role"
    if ($MOL) {
        return $true
    }
    return $false
}

Function DoesJobHaveMOLLogin ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $MOL = $DetailsOfJobTitle."MOL Role"
    if ($MOL) {
        if ($MOL -match "Shared Account") {
            return $false
        }
        return $true #if user does have access, but not shared account.
    }
    return $false
}

Function DoesJobGetSAGE ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $SAGE = $DetailsOfJobTitle."Sage 100"
    if ($SAGE) {
        return $true
    }
    return $false
}

Function DoesJobGetDocuware ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $DW = $DetailsOfJobTitle."SG: DW Accounting"
    $DW1 = $DetailsOfJobTitle."SG: DW AP"
    $DW2 = $DetailsOfJobTitle."SG: DW AP Invoice Approvers"
    $DW3 = $DetailsOfJobTitle."SG: DW Customer Care"
    $DW4 = $DetailsOfJobTitle."SG: DW HR"
    $DW4 = $DetailsOfJobTitle."SG: DW SMT"
    if ($DW -or $DW1 -or $DW2 -or $DW3 -or $DW4 -or $DW5) {
        return $true
    }
    return $false
}

Function IsJobAManager ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $ManagerChange = $DetailsOfJobTitle."Manager Change"
    if ($ManagerChange) {
        return $true
    }
    return $false
}

#Returns a Boolean
Function IsUserAManager ($EmployeeID, $AccessMatrix) {
    $User = Get-AdUser -Filter {employeeID -eq $EmployeeID -and Enabled -eq $true} -Properties description | Select-Object description
    $Desc = $User."description"
    $Result = IsJobAManager $AccessMatrix $Desc
    Return $Result
}

Function EmployeeIDMatchesName ($EmployeeID, $Name) {
    $NamesEmpId = GetUsersEmployeeIDByName $Name
    if ($EmployeeID -eq $NamesEmpId) {
        return $true
    }
    return $false
}

Function DoesJobGetGA ($AccessMatrix, $JobTitle) {
    $JobTitleArray = $AccessMatrix."Job Titles" #Load the Column of Job Titles
    $IndexOfJobTitle = $JobTitleArray.IndexOf($JobTitle) #Gets the row of the current Job Title
    $DetailsOfJobTitle = $AccessMatrix[$IndexOfJobTitle] 
    $GA = $DetailsOfJobTitle."GA"
    if ($GA) {
        return $true
    }
    return $false
}