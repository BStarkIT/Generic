$Users = Get-ADUser -filter * -searchbase 'OU=CRN354-22,OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties SamAccountName | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
foreach ($User in $Users) {
    Write-Output $SAM
        Set-ADUser -Identity $SAM -clear "extensionattribute1"
        Set-ADUser -Identity $SAM -add @{"extensionattribute1"="CRN354-22"}
}
