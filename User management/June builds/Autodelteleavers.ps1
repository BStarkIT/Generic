# Autodelte leavers
# Author        Brian Stark
# Date          13/06/2022
# Proofed       
# Tested date   
# Version       1.00
# Purpose       Automation of deletion of leavers
# Useage        will be set as exe on schedule task
#
# Changes       1.00    13/06/22    BS  Inital copy

$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Auto\$Date.txt" -append
$DatetoDel = (Get-Date).ToString('dd MMMM yyyy')
$LeaversList = Get-ADUser -Filter * -SearchBase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties Name, description, DistinguishedName |
Select-Object Name, SamAccountName, Description, DistinguishedName | Where-Object { $_.Description -match "xx DELETE after DATE" }
$LeaversOlderThan1MonthList = ForEach ($UserToDelete in $LeaversList) {
    $UserInfo = Get-ADUser -identity $UserToDelete.SamAccountName -Properties * |
    Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = { $_.Description -replace '^.*-' } } 
    $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
    If ($UserDelDate -lt $DateToDel) {
        Get-Aduser $UserToDelete.SamAccountName -Properties * | select-object Name, SamAccountName, Description, DistinguishedName
    }
}
if ($null -eq $LeaversOlderThan1MonthList) {
    Write-Host "no Accounts due for deletion"
    break
}
if (@($LeaversOlderThan1MonthList).Count -gt 1) {
    try {
        ForEach ($UserToDelete in $LeaversOlderThan1MonthList) {
            Get-ADUser -Filter "SamAccountName -eq '$($UserToDelete.SamAccountName)'" | Select-Object -ExpandProperty 'DistinguishedName' | Remove-ADObject -Recursive -confirm:$false
            Write-Host $UserToDelete
            $DelADUserList = $($UserToDelete.Name)
            $DeleteADUserList += "$DelADUserList`r`n"  
        }
    }
    catch {
        Write-Host "error on $UserToDelete"
        break 
    }
    $DeletedUserList = ""
    $path = "\\scotcourts.local\Home\P"
    $folders = Get-ChildItem $path -Filter "*xx DELETE after DATE*" -directory | select-object Name, PSChildName | select-object PSChildName, @{n = 'Name'; e = { $_.Name -replace '^.*-' } } 
    ForEach ($folder in $folders) {
        Write-Host "Users $folder marked for deletion"
        $FolderDate = [System.DateTime]$folder.Name
        If ($FolderDate -lt $DateToDel) {
            try {
                if (test-path $path\$($folder.PSChildname)) {
                    Write-Host "deleting P drive...."
                    Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                    $DelUserList = $($folder.PSChildName) 
                    $DeletedUserList += "$DelUserList`r`n"
                    Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
                }
            }       
            catch {
                # this catch is new & untested. SHOULD rename all files & folders within the path, rename has been tested, the path part has not.
                # once renamed, 2nd deletion attept.
                Get-ChildItem -path $path\$($folder.PSChildName) -Recurse | ForEach-Object  -begin { $count = 1 }  -process { rename-item $_ -NewName "file$count.txt"; $count++ }
                if (test-path $path\$($folder.PSChildname)) {
                    Write-Host "Second pass of $($folder.PSChildname)"
                    Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                    $DelUserList = $($folder.PSChildName) 
                    $DeletedUserList += "$DelUserList`r`n"
                    Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
                }
            }
        }
    }
    Write-Host "The Leavers older than 1 month have been deleted."
    break
}