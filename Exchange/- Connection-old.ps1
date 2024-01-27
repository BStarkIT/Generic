$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAUEX03.scotcourts.local/powershell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
