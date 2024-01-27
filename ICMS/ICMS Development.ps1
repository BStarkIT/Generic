$email = read-Host 'email address of user to add'
$groups = Get-ADGroup -Filter "name -like 'ICMS Development*'" | Select-Object Name | Select-Object -ExpandProperty Name
ForEach ($group in $groups) {
    #Add-ADGroupMember -Identity $group -Members $email
    Write-Output "$email added to $group"
}