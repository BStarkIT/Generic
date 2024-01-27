$Users = Get-ADUser -filter * -searchbase 'OU=ResourceMailboxes,DC=scotcourts,DC=local' -Properties EmailAddress,SamAccountName | Select-Object EmailAddress,SamAccountName
foreach ($User in $Users) {
        #$SAM = $User -ireplace [regex]::Escape("@scotcourts.gov.uk"), ""
        Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute2"
        Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute2"="RES-ROOM"}
}
