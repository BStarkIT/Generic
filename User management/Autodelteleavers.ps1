<#
Autodelte leavers
.SYNOPSIS
This PowerShell script is to Automate the deletion of leavers

.NOTES
will be set as exe on schedule task
Proofed by <null>
Approved on CAB CRN302-22
Automation approved on CAB 

This script performs the following on a User account.
Deletion of account in AD if in the disabled leavers OU.
Deletion of P drive
Account detection based on description "xx DELETE after DATE -- <Date>"
Contractors:
if the AD account is identified as Contractor (extensionattribute4 set to "Contractor") & has been disabled for 2 months due to account expiry, but has not been requested to be treated as a leaver, the leaver process is actioned anyway.

1.01    15/11/22    BS  Contractor subsection completed
1.00    13/06/22    BS  Inital copy
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>

# Variables 
$path = "\\scotcourts.local\Home\P"
$expiredAccounts = Search-ADAccount -AccountDisabled -SearchBase 'OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' | Select-Object SamAccountName
$1monthgrace = (Get-Date).AddDays(-60).ToString('dd MMM yyyy') 
$DateDel = (Get-Date).AddDays(30).ToString('dd MMM yyyy') 
$LeaversOlderThan1MonthList = @()
$Date = Get-Date -Format "dd-MM-yyyy"
$DatetoDel = (Get-Date).ToString('dd MMM yyyy')
$LeaversList = Get-ADUser -Filter * -SearchBase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties Name, description, DistinguishedName | Select-Object Name, SamAccountName, Description, DistinguishedName | Where-Object { $_.Description -match "xx DELETE after DATE --" }
# Script begin
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Auto\$Date.txt" -append
ForEach ($UserToDelete in $LeaversList) {
    $UserInfo = Get-ADUser -identity $UserToDelete.SamAccountName -Properties * |
    Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = { $_.Description -replace '^.*--' } } 
    $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
    If ($UserDelDate -lt $DateToDel) {
        $LeaversOlderThan1MonthList += Get-Aduser $UserToDelete.SamAccountName -Properties * | select-object SamAccountName
        Write-output $UserInfo.SamAccountName 
    }
}
# Contractor finder
Write-host "Disabled accounts Found"
Foreach ($expiredAccount in $expiredAccounts) {
    $Accountcheck = Get-ADUser $expiredAccount.SamAccountName -Properties * | Select-Object SamAccountName, extensionattribute4, Description, AccountExpirationDate
    Write-host "Checking account"
    if ($Accountcheck.extensionattribute4 -eq "Contractor") {
        Write-output "Contractor found"
        if ($Accountcheck.AccountExpirationDate -lt $1monthgrace) {
            $Excontractor = $Accountcheck.SamAccountName
            Write-host 'Contractor $Excontractor disabled for over 2 months, begining leaver process'
            $Usermembership = get-aduser -identity $Excontractor -property MemberOf | Foreach-Object { $_.MemberOf | Get-AdGroup | Select-Object Name, SamaccountName | Select-object -ExpandProperty SamAccountName }
            $pdrive = test-path \\scotcourts.local\home\P\$Excontractor
            try {
                If (($null -ne $Usermembership) -and ($pdrive -eq $true)) {                   
                    $Usermembership | out-file \\scotcourts.local\home\P\$Excontractor\UserMembershipBackup.csv
                    Write-Host "Backed up Users Security group and Distribution List membership to $Excontractor P drive." 
                }
            }
            catch {
                Write-Host "Error - Cant Users Security group and Distribution List membership to $Excontractor P drive." 
            }
            try {
                if (test-path \\scotcourts.local\Home\P\$Excontractor) {
                    Get-Item \\scotcourts.local\home\P\$Excontractor | Rename-Item -NewName { $_.Name -replace "$", " xx DELETE after DATE - $Datedel" }
                    Write-Host "Renamed $Excontractor P drive folder on SAUFS01" 
                }
            }
            catch {
                Write-Host "Error - Cannot Renamed $Excontractor P drive folder on SAUFS01" 
            }
            Get-ADuser -Identity $Excontractor -property msExchHideFromAddressLists | Set-ADObject -Replace @{msExchHideFromAddressLists = $true }
            Set-ADUser $Excontractor.SamAccountName -Description "$($_.Description) xx DELETE after DATE -- $Datedel"
            Get-ADUser $Excontractor.SamAccountName | Move-ADObject -targetpath 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
        }
    }
}
if ($null -eq $LeaversOlderThan1MonthList) {
    Write-Host "no Accounts due for deletion"
    break
}
else {
    Write-Host "begining deletion"
    try {
        ForEach ($User in $LeaversOlderThan1MonthList.SamAccountName) {
            Write-Host $User
            Get-ADUser $User | Set-ADObject -ProtectedFromAccidentalDeletion:$false -PassThru | Remove-ADObject -Recursive -confirm:$false
        }
    }
    catch {
        Write-Host "error on $User"
        Write-Host $User
        Write-host $error
        break
    }
    $DeletedUserList = ""
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
                    $DelUserList = $($folder.PSChildName) 
                    $DeletedUserList += "$DelUserList`r`n"
                    Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
                }
            }       
            catch {
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
    Pause
    break
}