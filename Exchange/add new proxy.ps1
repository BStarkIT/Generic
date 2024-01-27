$users = Get-ADUser -Filter "ProxyAddresses -like '*@scotcourts.pnn.gov.uk'" -SearchBase "DC=scotcourts,DC=local"
foreach ($user in $users)

{

$email = $user.samaccountname + '@scotcourtstribunals.pnn.gov.uk'

$newemail = "smtp:"+$email

Set-ADUser $user -Add @{proxyAddresses = ($newemail)}

}