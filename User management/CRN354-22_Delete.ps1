# Delete unused user accounts - CAB CRN354-22
# Author        Brian Stark
# Date          06/10/2022
# Proofed       
# Tested date   
# Version       1.00
# Purpose       CRN354-22 - delete accounts 
# Useage        Admin to delete accounts after 26 April 2023
#
$Users = Get-ADUser -Filter * -SearchBase 'OU=CRN354-22,OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties * | Select-Object SamAccountName, Description | Where-Object { $_.Description -match "xx DELETE after DATE" }
$DatetoDel = (Get-Date).ToString('dd MMMM yyyy')
$path = "\\scotcourts.local\Home\P"
ForEach ($User in $Users) {
    $UserInfo = Get-ADUser -identity $User.SamAccountName -Properties * |
    Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = { $_.Description -replace '^.*--' } } 
    $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
    If ($UserDelDate -lt $DateToDel) {
        Get-ADUser -Filter "SamAccountName -eq '$($UserInfo.SamAccountName)'" | Select-Object -ExpandProperty 'DistinguishedName' | Remove-ADObject -Recursive -confirm:$false
    }
    else {
        Write-host "not time for deletion"
    }
}
Write-Host "Searching for all P Drives, this may take some time"
$folders = Get-ChildItem $path -Filter "*xx DELETE after DATE*" -directory | select-object Name, PSChildName | select-object PSChildName, @{n = 'Name'; e = { $_.Name -replace '^.*-' } } 
ForEach ($folder in $folders) {
    Write-Host "Users $folder marked for deletion"
    $FolderDate = [System.DateTime]$folder.Name
    If ($FolderDate -lt $DateToDel) {
        try {
            if (test-path $path\$($folder.PSChildname)) {
                Write-Host "deleting P drive...."
                Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
            }
        }       
        catch {
            Get-ChildItem -path $path\$($folder.PSChildName) -Recurse | ForEach-Object  -begin { $count = 1 }  -process { rename-item $_ -NewName "file$count.txt"; $count++ }
            if (test-path $path\$($folder.PSChildname)) {
                Write-Host "Second pass of $($folder.PSChildname)"
                Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
            }
        }
    }
}
Write-Host "The Accounts disabled on CRN354-22 have been deleted."