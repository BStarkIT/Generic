<#
.SYNOPSIS
This PowerShell script is to create svc accounts with & without mailboxes.

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
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\svcmaker\$Date.txt" -append
$TempSharedMailbox = read-Host 'name of service account'
$TicketTrace = read-Host 'Ticket number'
$Mail = Read-Host 'Mail enabled? (y/n)'
$Description = Read-Host 'Description'
$SharedMailbox = $TempSharedMailbox -replace '[^\x30-\x39\x2D\x41-\x5A\x61-\x7A]+', ''
$Ticket = $TicketTrace -replace '[^\x30-\x39]+', ''
write-host $TicketTrace
Write-host $Ticket
$charlist1 = [char]97..[char]122
$charlist2 = [char]65..[char]90
$charlist3 = [char]48..[char]57
$charlist4 = [char]33..[char]38 + [char]40..[char]43 + [char]45..[char]46 + [char]64
$pwdList = @()
$pwLength = 6 
For ($i = 0; $i -lt $pwlength; $i++) {
    $pwdList += $charlist1 | Get-Random
    $pwdList += $charlist2 | Get-Random
    $pwdList += $charlist3 | Get-Random
    $pwdList += $charlist4 | Get-Random
}
$pass = -join ($pwdList | get-random -count $pwdList.count)
$password = ConvertTo-SecureString $pass -AsPlainText -Force
Write-Host "Please Note: This account will be made with Password: " $pass -ForegroundColor Red
$DisplayName = $SharedMailbox
$SAM = $SharedMailbox
$UPN = $SharedMailbox + "@scotcourts.gov.uk"
$mail = $SharedMailbox + "@scotcourts.gov.uk"
$Routingaddress = $SharedMailbox + "@scotcourtsgovuk.mail.onmicrosoft.com"
$OU = "OU=Service Accounts,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local"
$DC = "SAU-DC-04.scotcourts.local"
if ($Ticket -eq '') {
    Write-Host "blank ticket number"
    break
}
if ($Mail -eq "n") {
    New-AdUser -Name $DisplayName -SamAccountName $SAM -DisplayName $DisplayName -UserPrincipalName $UPN -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server $DC -AccountPassword $password -passThru
    Start-Sleep -Seconds 5
    Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "SVC-APP" }
    Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
    Set-ADUser -Identity $SAM -Description $Description -PasswordNeverExpires $true
    Write-Output "Service account with Mailbox $Sam Crated"   
    $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
    Pause
}
else {
    $DupCatch = Get-ADObject -Properties mail, proxyAddresses -Filter { mail -eq $mail -or proxyAddresses -eq "smtp:$mail" } 
    If ($null -eq $DupCatch) {
        New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -RemoteRoutingAddress $Routingaddress -DomainController $DC -Password $password -ResetPasswordOnNextLogon $false
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
        Set-ADUser -Identity $SAM -Description $Description -PasswordNeverExpires $true
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "SVC-APPM" }
        Set-ADUser -Identity $SAM -add @{"extensionattribute3" = "$Ticket" }
        Write-Output "Service account with Mailbox $Sam Crated"   
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Pause
    }
    else {
        Write-Output "Email address already in use."
        Pause
    }
}
