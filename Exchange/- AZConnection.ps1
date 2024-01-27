$UserCredential = Get-Credential
$Key = [byte]1..16
$UserCredential.Password | ConvertFrom-SecureString -Key $Key | Set-Content c:\PS\cred.key
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
Get-mailbox ajohnsontest@ScotCourts.gov.uk
Exit-PSSession
Remove-PSSession $Session
Write-Host "Exchange disconnected"
Start-Sleep -seconds 5
Invoke-Command -ComputerName  SAUAZADC01 -Credential $UserCredential -ScriptBlock {Start-ADSyncSyncCycle -PolicyType delta}