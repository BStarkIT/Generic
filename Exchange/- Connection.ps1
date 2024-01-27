$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://glw-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
