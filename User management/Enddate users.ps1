<#
export list of users with enddate in Ad
.SYNOPSIS
This PowerShell script is to add extensionattribute4 to contractors AD accounts made before extensionattribute4 was added to the newstart script to allow easier deletion.

.NOTES
Proofed by Vicky Pasieka & Gerry Kingham
Approved on CAB 

.DESCRIPTION
1.00    15/11/22    BS  Inital copy
#>
$Users = Get-ADUser -filter * -searchbase 'OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties * | Select-Object -ExpandProperty SamAccountName
foreach ($User in $Users) {
    $exp = Get-ADUser $User -Properties * | Select-Object -ExpandProperty AccountExpirationDate 
    if ($null -ne $exp) {
        $Result = get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate 
        if ($Result.description -like "*ontract*") {
            if ($Result.extensionattribute4 -notlike "Contractor") {
                Write-Output "Description only $User"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled  | Export-CSV C:\PS\Description_Only.csv -Append
            }
            Else {
                Write-Output "Description & Attrib 4 $User"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled  | Export-CSV C:\PS\Description_Attrib4.csv -Append
            }
        }
        else {
            if ($Result.extensionattribute4 -notlike "Contractor") {
                Write-Output "End date only $User"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled | Export-CSV C:\PS\Enddate_Only.csv -Append
            }
            Else {
                Write-Output "End date & Attrib 4 $User"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled | Export-CSV C:\PS\Enddate_Attrib4.csv -Append
            }
        }
    }
    else {
        $Result = get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4
        if ($Result.description -like "*ontract*") {
            if ($Result.extensionattribute4 -notlike "Contractor") {
                Write-Output "Description only $User - no end date"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled | Export-CSV C:\PS\Description_NoEnd.csv -Append

            }
            Else {
                Write-Output "Description & Attrib 4 $User - no end date"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled | Export-CSV C:\PS\Description_Attrib4_NoEnd.csv -Append
            }
        }
        else {
            if ($Result.extensionattribute4 -like "Contractor") {
                Write-Output "Attrib only $User - no end date"
                get-aduser $User -Properties * | Select-Object Description, SamAccountName, Office, Name, extensionattribute4, AccountExpirationDate, Enabled | Export-CSV C:\PS\Attrib4_noEnd.csv -Append
            }
        }
    }
}
    
