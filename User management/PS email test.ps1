$UserNameList = Get-aduser -Filter { SamAccountName -like "*Admin*" } -Properties DisplayName | Select-Object SamAccountName | Select-object -ExpandProperty SamAccountName
foreach ($UserName in $UserNameList) {
        $PDrive = test-path \\scotcourts.local\Home\P\$UserName
        Write-Output $PDrive
}