#This file is used to run individual scripts that make use of the functions in the Public Folder
Function UserManagement() {
    $ScriptPath = $PSScriptRoot +"\..\Full Scripts\UserManagement.ps1"
    &$ScriptPath #Execute Script
}