$Users = Get-ADComputer  -filter * -searchbase 'OU=Portable,OU=PCs,OU=SCTS,DC=scotcourts,DC=local' -Properties Name
foreach ($User in $Users) {
        Set-ADComputer -Identity $User -clear "extensionattribute2"
        Set-ADComputer -Identity $User -add @{"extensionattribute2"="DVC-LAPT"}
}
