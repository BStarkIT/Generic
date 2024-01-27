$Lists = import-csv -path "C:\PS\Ready\HRU.csv"
foreach ($User in $Lists ) {
    Write-Output $User.User
    $dn = $user.user
    $Sam = Get-ADUser -Filter { displayName -like $dn } | Select-Object -ExpandProperty SAMAccountName
    Write-Output "Found User $User"
    Write-output $User.Grade
    Set-ADUser $SAM -add @{"extensionattribute5" = $User.Grade}
    $Man = $User.Manager
    $Manager = Get-aduser -Filter {DisplayName -eq $Man} -Properties * | Select-Object DistinguishedName | Select-Object -ExpandProperty DistinguishedName 
    Write-Output "Found Manager $Manager"
    Set-ADUser $SAM -Manager $Manager -Department $User.Department -Title $User.Title
    Write-Output "Done"
}