#Script to check access in M365
#Connect-MsolService
$ExchangeAdminRoles = (Get-RoleGroup).Name
$date = (Get-Date).ToString("ddMMyyyy_hhmmss")
    $AzureAdminRoles = Get-MsolRole
    $UserEmailAddress = Read-Host -Prompt "Please enter email for user to search admin roles "
    Write-Host "Please wait while processing..." -ForegroundColor Green
    $ExchangeRoles = @()
    $AzureRoles = @()
    foreach($Role in $ExchangeAdminRoles)
        {
        if(Get-RoleGroupMember -Identity $Role | Where-Object  {$_.WindowsLiveID -like $UserEmailAddress } | Select-Object Name)
            {
            $ExchangeRoles += $Role
            }
        }

    foreach($Role in $AzureAdminRoles)
        {
        if(Get-MsolRoleMember -RoleObjectId $Role.objectid | Where-Object { $_.EmailAddress -eq $UserEmailAddress } )
            {
            $AzureRoles += $Role.Name
            }
        }

    if($AzureRoles.Count -ge $ExchangeRoles.Count){$Max = $AzureRoles.Count} else{$Max = $ExchangeRoles.Count}
    $i = 0
   
    While($i -lt $Max)
    {
    $Details = [pscustomobject]@{
        'EmailAddress' = $UserEmailAddress
        'AzureRoles' = $AzureRoles[$i]
        'ExchangeRoles' = $ExchangeRoles[$i]
        }
        $Details | Export-Csv -Path "C:\scripts\O365\reports\adminUserRoles_$($date).csv" -Append -NoTypeInformation
        $UserEmailAddress = ""
        $i++
    }

    $Roles = Import-Csv  -Path "c:\scripts\o365\reports\adminUserRoles_$($date).csv"
    $Roles
    Start-Sleep -Seconds 2

Write-Host "`n`rReport is generated at [Path : C:\scripts\O365\reports\adminUserRoles_$($date).csv ]" -ForegroundColor Green