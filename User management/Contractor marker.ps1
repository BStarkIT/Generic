<#
Cotractor Attrib4 adder
.SYNOPSIS
This PowerShell script is to add extensionattribute4 to contractors AD accounts made before extensionattribute4 was added to the newstart script to allow easier deletion.

.NOTES
Proofed by 
Approved on CAB 

Contractors:
if the AD account is identified as Contractor (extensionattribute4 set to "Contractor") & has been disabled for 2 months due to account expiry, but has not been requested to be treated as a leaver, the leaver process is actioned anyway.

1.00    15/11/22    BS  Inital copy
1.00    13/06/22    BS  Inital copy
Script written by Brian Stark 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$Users = $Users = Get-ADUser -filter * -searchbase 'OU=CRN354-22,OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties * | Select-Object Description, SamAccountName, extensionattribute4
foreach ($User in $Users) {
    $Name = $User.SamAccountName
    if ($User.description -like "Contract*") {
        if ($user.extensionattribute4 -notlike "Contractor") {
            Write-Output "extensionattribute4 would be added to $name"
            #Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute4"
            #Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute4" = "Contractor" }
        }
    }
    else {
        $exp = Get-ADUser $User.SamAccountName -Properties * | Select-Object -ExpandProperty AccountExpirationDate 
        if ($null -ne $exp) {
            Write-Output "unmarked Contractor found - $Name - $exp"
            $NewDescription = "Contractor - " + $User.Description 
            #Set-ADUser -Identity $User.SamAccountName -clear "extensionattribute4"
            #Set-ADUser -Identity $User.SamAccountName -add @{"extensionattribute4" = "Contractor" }
            #Set-ADUser -identity $User.SamAccountName -replace Description $NewDescription
        }
        else {
            Write-Output "User $name is staff"
        }

    }
}
    
