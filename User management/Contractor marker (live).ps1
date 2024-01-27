<#
Cotractor Attrib4 adder
.SYNOPSIS
This PowerShell script is to add extensionattribute4 to contractors AD accounts made before extensionattribute4 was added to the newstart script to allow easier deletion.

.NOTES
Proofed by Vicky Pasieka & Gerry Kingham
Approved on CAB 

.DESCRIPTION
1.00    15/11/22    BS  Inital copy
#>
$Users = $Users = Get-ADUser -filter * -searchbase 'OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties * | Select-Object Description, SamAccountName, extensionattribute4
foreach ($User in $Users) {
    $Name = $User.SamAccountName
    if ($User.description -like "*ontracto*") {
        if ($user.extensionattribute4 -notlike "nonfte") {
            Write-Output "extensionattribute4 being added to $name"
            Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute4"
            Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute4" = "nonfte" }
        }
    }
    else {
        $exp = Get-ADUser $User.SamAccountName -Properties * | Select-Object -ExpandProperty AccountExpirationDate 
        if ($null -ne $exp) {
            Write-Output "unmarked Contractor found - $Name - $exp"
            Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute4"
            Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute4" = "nonfte" }
        }
        else {
            Write-Output "User $name is staff"
        }

    }
}
