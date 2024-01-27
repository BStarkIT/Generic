$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauEX03.scotcourts.local/powershell -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking
