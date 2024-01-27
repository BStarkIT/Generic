$Tests = 'gkingtest2@scotcourts.gov.uk'
foreach ($Test in $Tests) {
    Write-Host $Test
    Set-Mailbox -Identity $Test -HiddenFromAddressListsEnabled $true
    Add-Content -Path C:\PS\Hiddenlog.txt -Value $Test
}
<#$Services = Get-ADUser -Filter 'name -like "SVC_*"' -Properties EmailAddress | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
foreach ($Service in $Services) {
    Write-host $service
    Set-Mailbox -Identity $Service -HiddenFromAddressListsEnabled $true
    Add-Content -Path C:\PS\Hiddenlog.txt -Value $Service
}
$Admins = Get-ADUser -Filter 'name -like "*_a*"' -Properties EmailAddress | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
foreach ($Admin in $Admins) {
    $Admin
    Set-Mailbox -Identity $Admin -HiddenFromAddressListsEnabled $true
    Add-Content -Path C:\PS\Hiddenlog.txt -Value $Admin
}
$RASS = Get-ADUser -Filter 'name -like "*_RAS"' -Properties EmailAddress | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
foreach ($RAS in $RASS) {
    $RAS
    Set-Mailbox -Identity $RAS -HiddenFromAddressListsEnabled $true
    Add-Content -Path C:\PS\Hiddenlog.txt -Value $RAS
}
#>