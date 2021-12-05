#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets eagle
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

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets MOL
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

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets an MOL Login, or the shared account. Shared account would be false.
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

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets Sage
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

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets Docuware
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

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job is defined as a manager in the access matrix
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

#Takes in custom ps object that is the access matrix
#Takes in an employee ID
#Returns a Boolean that states whether the users job title is defined as a manager
Function IsUserAManager ($EmployeeID, $AccessMatrix) {
    $User = Get-AdUser -Filter {employeeID -eq $EmployeeID -and Enabled -eq $true} -Properties description | Select-Object description
    $Desc = $User."description"
    $Result = IsJobAManager $AccessMatrix $Desc
    Return $Result
}

#Takes in an employee ID
#Takes in a username
#Returns a Boolean that states whether the emploee ID and the user name match.
Function EmployeeIDMatchesName ($EmployeeID, $Name) {
    $EmployeeID = [string]$EmployeeID
    $NamesEmpId = GetADUserCustom -SearchBy Name -SearchFor $Name -ReturnData EmployeeID
    if ($EmployeeID -eq $NamesEmpId) {
        return $true
    }
    return $false
}

#Takes in custom ps object that is the access matrix
#Takes in a job title
#Returns a Boolean that states whether the job gets Grower accounting.
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