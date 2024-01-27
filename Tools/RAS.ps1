<#
.SYNOPSIS
This PowerShell script is to create RAS accounts without mailboxes.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\RAS\$Date.txt" -append
Write-Host "Creating RAS account"
$tempfirstname = Read-Host 'First name'
$Templastname = Read-Host 'Surname'
$TicketTrace = read-Host 'Ticket number'
$Description = Read-Host 'Company'
$Ticket = $TicketTrace -replace '[^\x30-\x39]+', ''
$LastName = $Templastname -replace '[^\x2D\x41-\x5A\x61-\x7A]+', ''
$firstname = $tempfirstname -replace '[^\x2D\x41-\x5A\x61-\x7A]+', ''
$charlist1 = [char]97..[char]122
$charlist2 = [char]65..[char]90
$charlist3 = [char]48..[char]57
$charlist4 = [char]33..[char]38 + [char]40..[char]43 + [char]45..[char]46 + [char]64
$pwdList = @()
$pwLength = 2
For ($i = 0; $i -lt $pwlength; $i++) {
    $pwdList += $charlist1 | Get-Random
    $pwdList += $charlist2 | Get-Random
    $pwdList += $charlist3 | Get-Random
    $pwdList += $charlist4 | Get-Random
    $pwdList += $charlist1 | Get-Random
    $pwdList += $charlist2 | Get-Random
    $pwdList += $charlist3 | Get-Random
}
$pass = -join ($pwdList | get-random -count $pwdList.count)
$password = ConvertTo-SecureString $pass -AsPlainText -Force
Write-Host "Please Note: This account will be made with Password: " $pass -ForegroundColor Red
$TempDisplayName = $LastName + ", " + $FirstName + " - RAS"
$tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower() + "RAS"
$DisplayName = $TempDisplayName 
$samcatch = $tentativeSAM
$EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
    do {
        $incA ++
        $tentativeSAM = $samcatch + [string]$incA
        $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
    } 
    until (-not (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }))
}
if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
    do {
        $incB ++
        $TempDisplayName = $DisplayName + [string]$incB
    } 
    until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
}
$DisplayName = $TempDisplayName
$SAM = $tentativeSAM
Write-Host "RAS - $TempDisplayName Requested, Sam $SAM"
$OU = "OU=External Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
$UPN = $SAM + "@scotcourts.gov.uk"
$DC = "SAU-DC-04.scotcourts.local"
New-AdUser -GivenName $FirstName -Surname $Surname -Name $DisplayName -SamAccountName $SAM -DisplayName $DisplayName -UserPrincipalName $UPN -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server $DC -AccountPassword $password -passThru
Start-Sleep -Seconds 5
Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "RAS-USER" }
Set-ADUser -Identity $SAM -add @{"extensionattribute3" = "Created on Ticket : $Ticket" }
Set-ADUser -Identity $SAM -Description $Description -PasswordNeverExpires $true
Write-Output "RAS account $Sam Crated on Ticket $Ticket"   
$copy = "Username: $SAM  - Password: $pass" | clip
Pause

