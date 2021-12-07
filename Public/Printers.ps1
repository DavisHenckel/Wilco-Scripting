Function PrintADUsersCreatedInPastDays ($NumDays) {
    $DateCutOff=(Get-Date).AddDays(-$NumDays)
    Try {
        Get-ADUser -Filter * -Property whenCreated | Where-Object {$_.whenCreated -gt $datecutoff} | Format-Table Name, whenCreated -Autosize
    }
    Catch {
        Write-Host "Unable to run the command. Be sure to pass this function a day. IE to see past 10 days of created users, enter 10." -ForegroundColor Red
    }
}

#Simple function to print the countdown to open the file so the user isn't surprised when it pops up.
Function PrintTemplateCountdown {
    Write-Host "Opening Email Template in " -NoNewline -ForegroundColor Magenta
    Write-Host " 3" -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "2" -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "1" -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
    Start-Sleep -Milliseconds 333
    Write-Host "." -NoNewline -ForegroundColor Magenta
}

#basis function to interpret y or n
Function InterpretYesOrNo ($Prompt) {
    $Result = 'n'
    Write-Host -ForegroundColor Magenta $Prompt -NoNewLine
    While($true) {
        $Result = Read-Host
        if ($Result -eq "y" -or $Result -eq "Y") {
            return "y"
        }
        elseif ($Result -eq "N" -or $Result-eq "n") {
            return "n"
        }
        Write-Host "Invalid Input. Enter Y or N" -ForegroundColor Red
    }
} 