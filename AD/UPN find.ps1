$OrgUnit     = "DC=scotcourts,DC=local"
$UPNSuffix = 'SAUAZADC01'

Get-ADUser -Filter "userPrincipalName -like '*$UPNSuffix*'" -SearchBase $OrgUnit