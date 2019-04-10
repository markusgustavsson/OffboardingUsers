##  Yowza   
##  Just a quick little script written to facilitate employee offboarding. Check through comments or let me know if you need a run-through
##
Start-Transcript -Path "c:\temp\OffboardingScript $(get-date -f yyyy-MM-dd).txt"


$O365_Credentials  = Get-Credential 

##Create session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid" -Credential $O365_Credentials -Authentication Basic –AllowRedirection         
Import-PSSession $Session
Connect-MsolService -Credential $O365_Credentials

##Start by importing the activedirectory module
try{
    import-module activedirectory
    Write-Host -ForegroundColor Green "Importing AD cmdlets"
}
Catch{
    Write-host -ForegroundColor Red "Could not import activedirectory module. Please ensure it's installed."
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


$user = Read-Host -Prompt "Insert the name of the user you wish to offboard"



##set some variables
$employeedetails = get-aduser -filter { displayName -eq $user } -Properties *
$managerdetails = get-aduser $employeedetails.manager -Properties *
$DisabledUsersOU = 'OU PATH OF DISABLED USER IN AD'

$licenseRequired = "ORGANISATION:STANDARDWOFFPACK"  ## This is the license that's added to users to keep their mailbox active.
write-host "we are disabling $($employeedetails.samaccountname)" -ForegroundColor Green
write-host "the email that's being disabled is $($employeedetails.mail)" -ForegroundColor Green
Write-Host "The managers name is $($managerdetails.displayName)" -ForegroundColor Green
$confirmation = read-host -Prompt "Is this correct? Press n to terminate"

if ($confirmation -eq 'n') {
    Write-Host "terminating script"
    break
}


Write-Host "script not terminated, carry on"

$OutOfOfficeBody = @"
Hello <br>
Please mote that I am no longer working for The organisation <br>
"@

$Mailbox = get-mailbox -identity $employeedetails.mail
######################################  NO       ######################################
######################################  VARIABLES######################################
######################################  SET      ######################################
######################################  BELOW    ######################################


### Step 1 - Disable AD account 
try {
    Disable-ADAccount -identity $employeedetails.SamAccountName
    Write-Host "Account disabled in AD" -ForegroundColor Green
}
catch{
    Write-Host "Unable to disable account in AD. Pls check" -ForegroundColor Red
}
## Step 1.5 - Remove from GAL 
try {
    Set-ADUser -identity $employeedetails.SamAccountName -Replace @{msExchHideFromAddressLists=$true} 
    Write-Host "Removed from GAL " -ForegroundColor Green
}
catch {
    Write-Host "Unable to remove from GAL" -ForegroundColor Red
}
## Step 2 - Move to disabled accounts OU 
try{
    Get-ADUser $employeedetails.SamAccountName| Move-ADObject -TargetPath $DisabledUsersOU
    Write-Host "Account successfully moved" -ForegroundColor Green
}
catch {
    Write-Host "Unable to move to the disable OU - please check" -ForegroundColor Red
}


##### Disable OWA and other email services 
try{
    Set-CASMailbox -Identity $employeedetails.mail -OWAEnabled $false -PopEnabled $false -MAPIEnabled $false -ActiveSyncEnabled $false
    Write-Host "All mail apps disabled" -ForegroundColor Green
}
catch{
    Write-Host "Unable to disable email services for this account." -ForegroundColor Red
}

### Set out of office 
try{
    Set-MailboxAutoReplyConfiguration -Identity $employeedetails.mail -ExternalMessage $OutOfOfficeBody -InternalMessage $OutOfOfficeBody -AutoReplyState Enabled
    Write-Host "OOO message set up" -ForegroundColor Green
}
catch{
    Write-Host "Unable to set out of office" -ForegroundColor Red
}


#region Licenses
#### Remove licenses 
<# Script currently ensures that if a user has E2, it will remain. If there's an E3, it'll be changed to an E2. and all other licences removed. #>
Write-Host "removing licenses"
$GetLicenses = (Get-Msoluser -UserPrincipalName $employeedetails.mail).licenses.AccountSkuID
foreach($License in $GetLicenses){
    if ($License -eq $licenseRequired) {
        Write-Host "E2 license kept" -ForegroundColor Green
    }
    elseif ($License -eq "ORGANISATION:ENTERPRISEPACK") {
        Set-MsolUserLicense -UserPrincipalName $employeedetails.mail -RemoveLicenses "ORGANISATION:ENTERPRISEPACK" -AddLicenses $licenseRequired
        Write-Host "replaced E3 with E2" -ForegroundColor Green
    }
    else {
        Set-MsolUserLicense -UserPrincipalName $employeedetails.mail -RemoveLicenses $License
        Write-Host "removed license $($License)" -ForegroundColor Green
    }
}
#endregion


## Email manager to advise it's done. 
Write-Host "Emailing manager" -ForegroundColor Green


$Email_Username = 'EMAIL-ADDRESS' 
$encrypted = Get-Content -Path 'c:\scripts\password.txt'
$key = (1..16) 
$Email_Password = $encrypted | ConvertTo-SecureString -Key $key 
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Email_Username, $Email_Password 

$MailArgs = @{
    From       = $Email_Username
    To         = $managerdetails.mail
    Cc         = $Email_Username
    Subject    = "Account Details for $($employeedetails.displayname) "
    Body       = "Hi $($managerdetails.givenname),<br> 
The account of $($employeedetails.givenname) has been disabled as requested by HR. <br>
Please let us know if anything further is requested.<br>
Many thanks,<br>
IT"
SmtpServer = "smtp.office365.com"  ##outgoing mail server
Port       = 587
UseSsl     = $true
Credential = $Credentials
}
try{
    Send-MailMessage @MailArgs -BodyAsHtml
    Write-Host "Manager successfully emailed" -ForegroundColor Green
}
catch{
    Write-Host "email to manager failed" -ForegroundColor Red
}

Write-Host "Script completed - offboarding complete. Please review errors if any and let me know" -ForegroundColor Green

Remove-PSSession $Session
Stop-Transcript
