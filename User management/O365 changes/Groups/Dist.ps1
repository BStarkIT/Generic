$Users = Get-adgroup -filter {groupCategory -eq 'Distribution'} -searchbase 'OU=Distribution Lists,OU=Groups,OU=SCTS,DC=scotcourts,DC=local' -Properties SamAccountName
foreach ($User in $Users) {
        Set-adgroup -Identity $User -clear "extensionattribute2"
        Set-adgroup -Identity $User -add @{"extensionattribute2"="GRP-DIST"}
}

