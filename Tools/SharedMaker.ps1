<#
.SYNOPSIS
This PowerShell script is to create shared mailboxes.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>

$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Shared\$Date.txt" -append
$Proxies = @()
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$UserName = $env:username
if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
}
else {
    Write-host "* Please enter the following values - Name, Ticket number, Owner & Description *"
    Write-host "Email address will be auto generated"
    $TempSharedMailbox = read-Host 'Display name of Mailbox'
    $TicketTrace = read-Host 'Ticket number'
    $Owner2 = read-Host 'Owner'
    $description = Read-Host 'Description'
    $SharedMailbox = $TempSharedMailbox -replace '[^\x30-\x39\x2D\x41-\x5A\x61-\x7A]+', ''
    $Ticket = $TicketTrace -replace '[^\x30-\x39]+', ''
    $Owner1 = $Owner2 -replace '[^\x20\x2D\x41-\x5A\x61-\x7A]+', ''
    $DisplayName = $SharedMailbox
    $Mailbox = $SharedMailbox -replace " ", ""
    $Owner = "Owner: " + $Owner1
    $SAM = $Mailbox
    $UPN = $Mailbox + "@scotcourts.gov.uk"
    $mail = $Mailbox + "@scotcourts.gov.uk"
    $Routingaddress = $Mailbox + "@scotcourtsgovuk.mail.onmicrosoft.com"
    $OU = "OU=Shared Mailboxes,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local"
    $DC = "SAU-DC-04.scotcourts.local"
    $DupCatch = Get-ADObject -Properties mail, proxyAddresses -Filter { mail -eq $mail -or proxyAddresses -eq "smtp:$mail" } 
    if ($Ticket -eq '') {
        Write-Host "blank ticket number"
        break
    }
    If ($null -eq $DupCatch) {
        New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -RemoteRoutingAddress $Routingaddress -DomainController $DC -Shared
        Start-Sleep -Seconds 10
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
        Set-ADUser -Identity $SAM -Description $description -Office $owner
        Write-Output "Shared Mailbox $Sam created"   
        Pause
    }
    else {
        Write-Output "Email address already in use."
        Pause
    }
}
