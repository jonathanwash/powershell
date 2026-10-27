##Employee termination script: Disables, removes group memberships, hides from GAL, moves to correct OU##
##May be old in the days of MFA but the M365 parts converts mailbox to shared, adds manager's email to forward, removes M365 licenses, and syncs changes between AD and M365##

Write-Host "Capturing terminated employee and manager details"
$_user = Read-Host -Prompt "Please input terminated employee's Active Directory username:"
$_useremail = "$_user@domain.com"
$_manager = Read-Host "Please enter forwarding manager's username for email (first.last):"
$_mangeremail = "$_manager@domain.com"

#Disable user account
Disable-ADAccount -Identity $_user
Write-Host "Disabled user account"

#Remove user account from all group memberships
$ADgroups = Get-ADPrincipalGroupMembership -Identity  $_user | where {$_.Name -ne “Domain Users”}
Remove-ADPrincipalGroupMembership -Identity  $_user -MemberOf $ADgroups -Confirm:$False
Write-Host "Removed from all groups"
 
#Hide user account from Global Address List
Set-ADUser -Identity $_user -Replace @{msExchHideFromAddressLists=$True}
Write-Host "Hid from Global Address List"

#Move user account to disabled/shared users OU
Get-ADUser $_user | Move-ADObject -TargetPath 'OU=Shared,OU=Users,DC=domain,DC=com'
Write-Host "Moved to disabled OU"

#Capture Office 365 credentials
$AdminCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 

https://outlook.office365.com/powershell-liveid/ -Credential $AdminCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

Connect-MsolService -Credential $AdminCredential

Get-Mailbox -Identity $_useremail | Set-Mailbox -Type “Shared”
Write-Host "Mailbox set to shared"

Set-MsolUserLicense -UserPrincipalName $_useremail -RemoveLicenses domain:ENTERPRISEPACK
Set-MsolUserLicense -UserPrincipalName $_useremail -RemoveLicenses domain:O365_BUSINESS

(Get-MsolUser -UserPrincipalName $_useremail).licenses.AccountSkuId |
foreach{
    Set-MsolUserLicense -UserPrincipalName $_useremail -RemoveLicenses $_
}
Write-Host "Office 365 license(s) removed"

Set-Mailbox -Identity $_useremail -DeliverToMailboxAndForward $True -ForwardingSMTPAddress $_mangeremail
Write-Host "Manager's email address added to forwarding"

Invoke-Command -ComputerName ServerWithADSyncInstalled -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

Write-Host "If nothing is in red everything worked; else read the error"
Write-Host -NoNewLine "Press any key to continue..."
$Host.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL
$Host.UI.RawUI.Flushinputbuffer()
