<#
.SYNOPSIS
This PowerShell script is to create shared mailboxes.
Must be run as local admin.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$SharedMailbox = read-Host 'name of Mailbox'
$Ticket = read-Host 'Ticket number'
$Owner = read-Host 'Owner Name'
$Description = Read-Host 'Description'
$Office = "Owner: " + $Owner
$DisplayName = $SharedMailbox
$SAM = $SharedMailbox
$UPN = $SharedMailbox + "@scotcourts.gov.uk"
$mail = $SharedMailbox + "@scotcourts.gov.uk"
$Routingaddress = $SharedMailbox + "@scotcourtsgovuk.mail.onmicrosoft.com"
$OU = "OU=Shared Mailboxes,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local"
$DC = "SAU-DC-04.scotcourts.local"
$DupCatch = Get-ADObject -Properties mail, proxyAddresses -Filter {mail -eq $mail -or proxyAddresses -eq "smtp:$mail"} 
If ($null -eq $DupCatch){
    New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -DomainController $DC -Shared
    Start-Sleep -Seconds 5
    $TribProxy = "smtp:$SAM@scotcourtstribunals.gov.uk"
    $SCSPrimary = "SMTP:$SAM@scotcourts.gov.uk"
    $newProxy1 = "smtp:$SAM@scotcourts.pnn.gov.uk"
    $newProxy2 = "smtp:$SAM@scotcourtstribunals.pnn.gov.uk"
    $newProxy3 = "smtp:$SAM@scotcourtsgovuk.mail.onmicrosoft.com"
    $newProxy4 = "X400:C=GB;A=CWMAIL;P=SCS;O=SCOTTISH COURTS;S=" + $SAM + ";"
    $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
    foreach ($Proxy in $Proxies) {
        Set-ADUser -identity $SAM -add @{proxyAddresses = ($Proxy) }
    }
    Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "MBS-SHAR" }
    Set-ADUser -Identity $SAM -add @{"extensionattribute3" = "$Ticket" }
    Set-ADUser -Identity $SAM -Office $Office -Description $Description 
    Write-Output "Shared Mailbox $Sam Crated"   
    Pause
}
else {
    Write-Output "Email address already in use."
    Pause
}