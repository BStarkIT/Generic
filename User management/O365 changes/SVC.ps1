$Users = Get-ADUser -filter * -searchbase 'ou=Service Accounts,ou=Resource Accounts,ou=scts,DC=scotcourts,DC=local' -Properties SamAccountName
foreach ($User in $Users) {
        Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute2"
        Set-ADUser -Identity $User -add @{"extensionattribute2"="SVC-APPM"}
}
$Users = Get-ADUser -filter * -searchbase 'OU=Service Accounts,DC=scotcourts,DC=local' -Properties SamAccountName
foreach ($User in $Users) {
        Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute2"
        Set-ADUser -Identity $User -add @{"extensionattribute2"="SVC-APPM"}
}
