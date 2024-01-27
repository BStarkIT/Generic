<#
Delete User Accounts for Leavers
Version 5.01
13/07/2022
.SYNOPSIS
This PowerShell script is to process the accounts of Leavers.

.NOTES
helpdesk staff run a shorcut to Users - Disable Leaver Accounts.
13/07/2022  BS  5.0 Rebuild of script from scratch matching current requirements
13/06/2022  BS  4.0 O365 changes
02/09/2020  BS  2.1 Update & clean up of script & fix issues.
09/01/2020  GK  1.2 Changes to take acount of W10 OU users. Improves readability.
05/12/2019  GK  2.03 Replace references to \\scotcourts.local\data\users with \\scotcourts.local\home\p. Replace references to changed string with variable 
06/12/2018  JM  1. user accounts with blackberries won't delete with "remove-aduser" changed to "remove-adobject" & added DeleteUserADAccount function.
                2. some user profile folders have filenames longer than 255 characters changed Remove-Item -path to "\\?\$path" in DeleteProfileFolders function.    

This script performs the following on a User account.
Deletion of account in AD if in the disabled leavers OU.
Deletion of P drive
Account detection based on description ""

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>


$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Delete\$Date.txt" -append
$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\Delete.ico'
$message = "

This script searches the Z-Disabled_Leavers OU in AD for User accounts marked for deletion.

If the 'delete after date' has passed the Users account will be:

AD: Users AD account will be deleted.

SAUFS01: Users P drive folder on \\scotcourts.local\home\p will be deleted."
#
# Start of script:
#
$version = '5.01'
#  Show Start Message:   #
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", 'Delete User Accounts for Leavers.', 
    [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {
    Exit
} 
else {
    $UserName = $env:username
    if ($UserName -notlike "*stark") {
        Write-Host "Must be run as Admin, Script run as $UserName"
        Pause
    }
    else {
        # create the pop up information form   #
        Function PopUpForm {
            Add-Type -AssemblyName System.Windows.Forms    
            # 
            $PopForm = New-Object System.Windows.Forms.Form
            $PopForm.Text = "Delete User Accounts for Leavers."
            $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
            $PopForm.Size = New-Object System.Drawing.Size(420, 200)
            $PopUpLabel = New-Object System.Windows.Forms.Label
            $PopUpLabel.Location = '80, 40' 
            $PopUpLabel.Size = '300, 80'
            $PopUpLabel.Text = $poplabel
            $PopForm.Controls.Add($PopUpLabel)
            $PopForm.Show() | Out-Null
            Start-Sleep -Seconds 2
            $PopForm.Close() | Out-Null
        }
        #
        # create the pop up information form complete   #
        #
        # create the Select Multiple User AD account function   #
        Function DeleteMultipleUserAccount {
            # get list of users with name and description to display in form   #
            $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) { $($User.Name) + " - " + $($User.Description) }
            # create main form   #
            [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
            $DeletingUsersForm = New-Object System.Windows.Forms.Form
            $DeletingUsersForm.width = 745
            $DeletingUsersForm.height = 495
            $DeletingUsersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
            $DeletingUsersForm.Controlbox = $false
            $DeletingUsersForm.Icon = $Icon
            $DeletingUsersForm.FormBorderStyle = 'Fixed3D'
            $DeletingUsersForm.Text = "Delete Leavers User Accounts v$version."
            $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
            $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox1.Location = '40,20'
            $LeaverBox1.size = '650,90'
            $LeaverBox1.text = '1. Check: This deletes Users AD accounts and mailboxes:'
            $LeavertextLabel1 = New-Object System.Windows.Forms.Label
            $LeavertextLabel1.Location = '20,25'
            $LeavertextLabel1.size = '600,30'
            $LeavertextLabel1.Text = 'Check the Users in the list below to ensure they are all marked to be deleted.' 
            $LeavertextLabel2 = New-Object System.Windows.Forms.Label
            $LeavertextLabel2.Location = '20,55'
            $LeavertextLabel2.size = '600,30'
            $LeavertextLabel2.Text = 'IF THERE ARE ANY USERS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
            $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox2.Location = '40,120'
            $LeaverBox2.size = '650,225'
            $LeaverBox2.text = '2. Check the Users below are labelled to be Deleted'
            $LeavertextLabel4 = New-Object System.Windows.Forms.ListBox    
            $LeavertextLabel4.Location = '40,20'
            $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
            $LeavertextLabel4.size = '570,170'
            $LeavertextLabel4.Datasource = $LeaversListToDisplayInForm 
            $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox3.Location = '40,355'
            $LeaverBox3.size = '650,30'
            $LeaverBox3.text = '3. Click Ok to DELETE User Accounts or Exit:'
            $LeaverBox3.button
            $OKButton = new-object System.Windows.Forms.Button
            $OKButton.Location = '590,395'
            $OKButton.Size = '100,40'          
            $OKButton.Text = 'Ok'
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = '470,395'
            $CancelButton.Size = '100,40'
            $CancelButton.Text = 'Exit'
            $CancelButton.add_Click( {
                    $DeletingUsersForm.Close()
                    $DeletingUsersForm.Dispose()
                    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
            $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
            $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
            $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
            $DeletingUsersForm.Add_Shown( { $DeletingUsersForm.Activate() })    
            $dialogResult = $DeletingUsersForm.ShowDialog()
            if ($dialogResult -eq 'OK') {
                Write-Host " Accepted DeleteLeavers form and starting to Remove User AD Account" 
                DeleteUserADAccount
            }
        }
        # create the Select Multiple User AD accounts function complete
        #
        # create the Select One User AD account function   #
        Function DeleteOneUserAccount {
            # get users with name and description to display in form   #
            $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) { $($User.Name) + " - " + $($User.Description) }
            [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
            [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
            $DeletingUsersForm = New-Object System.Windows.Forms.Form
            $DeletingUsersForm.width = 745
            $DeletingUsersForm.height = 495
            $DeletingUsersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
            $DeletingUsersForm.Controlbox = $false
            $DeletingUsersForm.Icon = $Icon
            $DeletingUsersForm.FormBorderStyle = 'Fixed3D'
            $DeletingUsersForm.Text = "Delete Leavers User Accounts v$Version."
            $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
            $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox1.Location = '40,20'
            $LeaverBox1.size = '650,90'
            $LeaverBox1.text = '1. Check: This deletes Users AD accounts and mailboxes:'
            $LeavertextLabel1 = New-Object System.Windows.Forms.Label
            $LeavertextLabel1.Location = '20,25'
            $LeavertextLabel1.size = '600,30'
            $LeavertextLabel1.Text = 'Check the User in the list below to ensure they are marked to be deleted.' 
            $LeavertextLabel2 = New-Object System.Windows.Forms.Label
            $LeavertextLabel2.Location = '20,55'
            $LeavertextLabel2.size = '600,30'
            $LeavertextLabel2.Text = 'IF THE ANY USER IS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
            $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox2.Location = '40,120'
            $LeaverBox2.size = '650,225'
            $LeaverBox2.text = '2. Check the User below is labelled to be Deleted'
            $LeavertextLabel4 = New-Object System.Windows.Forms.Label    
            $LeavertextLabel4.Location = '40,20'
            $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
            $LeavertextLabel4.size = '570,170'
            $LeavertextLabel4.Text = $LeaversListToDisplayInForm 
            $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
            $LeaverBox3.Location = '40,355'
            $LeaverBox3.size = '650,30'
            $LeaverBox3.text = '3. Click Ok to DELETE the User Account or Exit:'
            $LeaverBox3.button
            $OKButton = new-object System.Windows.Forms.Button
            $OKButton.Location = '590,395'
            $OKButton.Size = '100,40'          
            $OKButton.Text = 'Ok'
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = '470,395'
            $CancelButton.Size = '100,40'
            $CancelButton.Text = 'Exit'
            $CancelButton.add_Click( {
                    Write-Host "Canceled"
                    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
            $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
            $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
            $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
            $DeletingUsersForm.Add_Shown( { $DeletingUsersForm.Activate() })    
            $dialogResult = $DeletingUsersForm.ShowDialog()
            if ($dialogResult -eq 'OK') {
                Write-Host " Accepted DeleteLeavers form and starting to Remove User AD Account" 
                DeleteUserADAccount
            }
        }
        # create the Select one User AD accounts function complete   #      
        #
        # create the Delete User AD account function   #
        Function DeleteUserADAccount {     
            # create a list to add the deleted user names to   #
            $DeleteADUserList = ""
            try {
                ForEach ($UserToDelete in $LeaversOlderThan1MonthList) {
                    Get-ADUser -Filter "SamAccountName -eq '$($UserToDelete.SamAccountName)'" | Select-Object -ExpandProperty 'DistinguishedName' | Remove-ADObject -Recursive -confirm:$false
                    # original remove aduser to be removed  #
                    Write-Host $UserToDelete
                    $DelADUserList = $($UserToDelete.Name)
                    $DeleteADUserList += "$DelADUserList`r`n"  
                    $poplabel = "Deleting the AD account for User`n`n$($UserToDelete.Name)."
                    PopupForm
                }
            }
            catch {
                Write-Host "error on $UserToDelete"
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG deleting the User`n`n$($UserToDelete.Name)`n`nAD account !!!.`n`nPlease contact 3rd line with the details.", 'Delete User Accounts for Leavers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $DeletingUsersForm.Close()
                break 
            }
            DeleteUserFolders
        }
        # Delete User AD account function complete   #
        #
        # create the Delete Users Folders function   #
        Function DeleteUserFolders {  
            # Get User Folders with Delete after Date more than 1 month old   #
            $DeletedUserList = ""
            $path = "\\scotcourts.local\Home\P"
            Write-Host "Searching for all P Drives, this may take some time"
            $folders = Get-ChildItem $path -Filter "*xx DELETE after DATE*" -directory | select-object Name, PSChildName | select-object PSChildName, @{n = 'Name'; e = { $_.Name -replace '^.*-' } } 
            ForEach ($folder in $folders) {
                Write-Host "Users $folder marked for deletion, Checking date"
                $FolderDate = [System.DateTime]$folder.Name
                If ($FolderDate -lt $DateToDel) {
                    # delete user folder on scotcourts.local\Home\P   #
                    try {
                        if (test-path $path\$($folder.PSChildname)) {
                            Write-Host "deleting P drive...."
                            Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                            $DelUserList = $($folder.PSChildName) 
                            $DeletedUserList += "$DelUserList`r`n"
                            Write-Host "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." 
                            $poplabel = "Deleting the User`n`n$($folder.PSChildname)`n`nP drive folder on $path."
                            PopupForm
                        }
                    }       
                    catch {
                        Write-Host "error deleting $folder.PSChildname"
                        [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG removing the User`n`n$($folder.PSChildname)`n`nP drive folder !!!.`n`nPlease contact the 3rd line with the details.", 'Delete User Accounts for Leavers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        # remove variables #
                        Remove-Variable "UserToDelete", "UserInfo", "DelADUserList", "DeleteADUserList"
                        $DeletingUsersForm.Close()
                        break  
                    }
                }
            }
            Write-Host "The Leavers older than 1 month have been deleted." 
            [System.Windows.Forms.MessageBox]::Show("The Delete Leavers process is complete.`n`nA list of deleted user accounts will be emailed.", 'Delete leavers.',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            break
        }
        # create the Delete Users Folders function complete   #
        #
        #
        # create the main form   #
        #
        # Set Delete Date to -1 month   #
        $DatetoDel = (Get-Date).ToString('dd MMMM yyyy')
        # get list of users in the Z-Disabled_Leavers OU that are marked to be deleted   #
        $LeaversList = Get-ADUser -Filter * -SearchBase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties Name, description, DistinguishedName |
        Select-Object Name, SamAccountName, Description, DistinguishedName | Where-Object { $_.Description -match "xx DELETE after DATE" }
        # get users older than 1 month   #
        $LeaversOlderThan1MonthList = ForEach ($UserToDelete in $LeaversList) {
            # get date on users AD description   #
            $UserInfo = Get-ADUser -identity $UserToDelete.SamAccountName -Properties * |
            Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = { $_.Description -replace '^.*-' } } 
            $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
            # check if users delete date is over 1 month   #
            If ($UserDelDate -lt $DateToDel) {
                Get-Aduser $UserToDelete.SamAccountName -Properties * | select-object Name, SamAccountName, Description, DistinguishedName
            }
        }
        # if no users to delete show message and exit   #
        if ($null -eq $LeaversOlderThan1MonthList) {
            Write-Host "no Accounts due for deletion"
            [System.Windows.Forms.MessageBox]::Show("There are currently no Leavers to delete.`n`nThe Delete Leavers process is complete.", 'Delete Leavers.',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            break
        }
        if (@($LeaversOlderThan1MonthList).Count -eq 1) {
            DeleteOneUserAccount
        }
        Else {
            DeleteMultipleUserAccount
        }
    }
}
