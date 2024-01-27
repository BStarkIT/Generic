$Users = Get-ADUser -filter * -searchbase 'OU=User Accounts (Admin),OU=SCTS,DC=scotcourts,DC=local' -Properties SamAccountName
foreach ($User in $Users) {
        Set-ADUser -Identity $User -clear "extensionattribute2"
        Set-ADUser -Identity $User -add @{"extensionattribute2"="ADM-PERS"}
}
