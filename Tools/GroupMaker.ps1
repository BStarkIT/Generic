<#
.SYNOPSIS
This PowerShell script is to create Security or Distribution groups.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session

$GroupName1 = read-Host 'name of group'
$GroupName = $GroupName1 -replace '[^\x30-\x39\x2D\x41-\x5A\x61-\x7A]+', ''
$Ticket1 = read-Host 'Ticket number'
$Ticket = $Ticket1 -replace '[^\x30-\x39]+', ''
$Mail = Read-Host 'Mail enabled? (y/n)'
$Description = Read-Host 'Description'

$DisplayName = $GroupName
$SAM = $GroupName
$UPN = $GroupName + "@scotcourts.gov.uk"
$mail = $GroupName + "@scotcourts.gov.uk"
$Routingaddress = $GroupName + "@scotcourtsgovuk.mail.onmicrosoft.com"

$DC = "SAU-DC-04.scotcourts.local"
if ($Mail -eq "n") {
    $OU = "OU=Service Accounts,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local"
    
    Pause
}
else {
    $DupCatch = Get-ADObject -Properties mail, proxyAddresses -Filter { mail -eq $mail -or proxyAddresses -eq "smtp:$mail" } 
    If ($null -eq $DupCatch) {
       
        Pause
    }
    else {
        Write-Output "Email address already in use."
        Pause
    }
}
