$US = read-Host 'Name of User to be located'
$user = Get-ADUser -Identity $US -Properties CanonicalName, LastLogonTimeStamp
$userOU = ($user.DistinguishedName -split ",", 3)[-1]
Write-Output $userOU