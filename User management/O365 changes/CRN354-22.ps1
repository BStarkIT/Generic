$Users = Get-ADUser -filter * -searchbase 'OU=CRN354-22,OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties EmailAddress | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
foreach ($User in $Users) {
    Write-Output $SAM
        Set-ADUser -Identity $SAM -clear "extensionattribute1"
        Set-ADUser -Identity $SAM -add @{"extensionattribute"="CRN354-22"}
}