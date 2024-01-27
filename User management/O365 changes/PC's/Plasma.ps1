$Users = Get-ADComputer  -filter * -searchbase 'OU=Desktop,OU=PCs (Bespoke Role),OU=SCTS,DC=scotcourts,DC=local' -Properties Name
foreach ($User in $Users) {
        Set-ADComputer -Identity $User -clear "extensionattribute2"
        Set-ADComputer -Identity $User -add @{"extensionattribute2"="DVC-PLAS"}
}
