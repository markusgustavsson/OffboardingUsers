##  Yowza   
##  Just a quick little script written to facilitate employee offboarding. Check through comments or let me know if you need a run-through
##
Start-Transcript -Path c:\temp\OffboardingScript $(get-date -f yyyy-MM-dd).txt

##Start by importing the activedirectory module
try{
    import-module activedirectory
    Write-Host -ForegroundColor Green "Importing AD cmdlets"
}
Catch{
    Write-host -ForegroundColor Green "Could not import activedirectory module. Please ensure it's installed."
}

try{
    Import-module MSOnline
    Write-host -ForegroundColor Green "Importing MsOnline cmdlets"
}
Catch{
    Write-Host -ForegroundColor Red "Could not load MsOnline module. Ensure it is installed."
    add-AlertDialog -title "Importing MsOnline cmdlets" -message "Could not load MsOnline module. Ensure it is installed"
    Exit
}

## connect to O365 

#get credentials for O365
$Office365Credentials  = Get-Credential 

##Create session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $Office365credentials -Authentication Basic –AllowRedirection         

##import session
Import-PSSession $Session
Import-Module MSonline
Connect-MsolService -Credential $Office365Credentials


$offboardUser = Read-Host -Prompt "Insert the name of the user you wish to offboard"

Get-ADUser -Filter { displayName -like $user } | select name, samaccountname
