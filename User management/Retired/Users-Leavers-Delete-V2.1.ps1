# Delete User Accounts for Leavers v2.0
#region Info
# Author        Brian Stark
# Date          15/07/2020
# Version       2.0
# Purpose       To disable accounts of leavers
# Useage        helpdesk staff run a shorcut to Delete User Accounts for Leavers v1.1.exe
#               
# Revisions     
#               V2.0  19/06/2020 BS - Update to current standards & include all current OU's
#               V1.2  09/01/2020 GK - Changes to take acount of W10 OU users. Improves readability.
#               V2.3  05/12/2019 GK - Replace references to \\scotcourts.local\data\users with \\scotcourts.local\home\p. Replace references to changed string with variable 
#               V1.1  06/12/2018 JM - 1. user accounts with blackberries won't delete with "remove-aduser" changed to "remove-adobject" & added DeleteUserADAccount function.
#                               2. some user profile folders have filenames longer than 255 characters changed Remove-Item -path to "\\?\$path" in DeleteProfileFolders function.    
#  
#endregion Info
# Script function: This script performs the following on a User account.
$message = "

This script searches the Z-Disabled_Leavers OU in AD for User accounts marked for deletion.

If the 'delete after date' has passed the Users account will be:

AD: Users AD account will be deleted.

SAUFS01: Users profile folder on \\scotcourts.local\data\profiles will be deleted.

SAUFS01: Users P drive folder on \\scotcourts.local\home\p will be deleted."
#
# Start of script:
#
$version='2.0'
##  Show Start Message:   ###
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", 'Delete User Accounts for Leavers.', 
    [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {
    Exit
} 
else {
    ## create the pop up information form   ###
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        # create form
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = "Delete User Accounts for Leavers."
        $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $PopForm.Size = New-Object System.Drawing.Size(420, 200)
        # Add Label
        $PopUpLabel = New-Object System.Windows.Forms.Label
        $PopUpLabel.Location = '80, 40' 
        $PopUpLabel.Size = '300, 80'
        $PopUpLabel.Text = $poplabel
        $PopForm.Controls.Add($PopUpLabel)
        # Show the form
        $PopForm.Show()| Out-Null
        # wait 2 seconds
        Start-Sleep -Seconds 2
        # close form
        $PopForm.Close() | Out-Null
    }
    #
    ## create the pop up information form complete   ###
    #
    ## create the Select Multiple User AD account function   ###
    Function DeleteMultipleUserAccount {
        ## get list of users with name and description to display in form   ###
        $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) {$($User.Name) + " - " + $($User.Description)}
        ## create main form   ###
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $DeletingUsersForm = New-Object System.Windows.Forms.Form
        $DeletingUsersForm.width = 745
        $DeletingUsersForm.height = 495
        $DeletingUsersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $DeletingUsersForm.Controlbox = $false
        $DeletingUsersForm.Icon = $Icon
        $DeletingUsersForm.FormBorderStyle = 'Fixed3D'
        $DeletingUsersForm.Text = "Delete Leavers User Accounts v$version."
        $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox1.Location = '40,20'
        $LeaverBox1.size = '650,90'
        $LeaverBox1.text = '1. Check: This deletes Users AD accounts and mailboxes:'
        ### Create group 1 box text labels. ###
        $LeavertextLabel1 = New-Object System.Windows.Forms.Label
        $LeavertextLabel1.Location = '20,25'
        $LeavertextLabel1.size = '600,30'
        $LeavertextLabel1.Text = 'Check the Users in the list below to ensure they are all marked to be deleted.' 
        $LeavertextLabel2 = New-Object System.Windows.Forms.Label
        $LeavertextLabel2.Location = '20,55'
        $LeavertextLabel2.size = '600,30'
        $LeavertextLabel2.Text = 'IF THERE ARE ANY USERS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
        ### Create group 2 box in form. ###
        $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox2.Location = '40,120'
        $LeaverBox2.size = '650,225'
        $LeaverBox2.text = '2. Check the Users below are labelled to be Deleted'
        # Create group 2 box text labels.
        $LeavertextLabel4 = New-Object System.Windows.Forms.ListBox    
        $LeavertextLabel4.Location = '40,20'
        $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $LeavertextLabel4.size = '570,170'
        $LeavertextLabel4.Datasource = $LeaversListToDisplayInForm 
        ### Create group 3 box in form. ###
        $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox3.Location = '40,355'
        $LeaverBox3.size = '650,30'
        $LeaverBox3.text = '3. Click Ok to DELETE User Accounts or Exit:'
        $LeaverBox3.button
        ### Add an OK button ###
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,395'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '470,395'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $DeletingUsersForm.Close()
                $DeletingUsersForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        ### Add all the Form controls ### 
        $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
        $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
        #### Activate the form ###
        $DeletingUsersForm.Add_Shown( {$DeletingUsersForm.Activate()})    
        #### Get the results from the button click ###
        $dialogResult = $DeletingUsersForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteLeavers form and starting to Remove User AD Account" -Verbose
            DeleteUserADAccount
        }
    }
    ## create the Select Multiple User AD accounts function complete   ###
    #
    ## create the Select One User AD account function   ###
    Function DeleteOneUserAccount {
        ## get users with name and description to display in form   ###
        $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) {$($User.Name) + " - " + $($User.Description)}
        ## create main form   ###
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $DeletingUsersForm = New-Object System.Windows.Forms.Form
        $DeletingUsersForm.width = 745
        $DeletingUsersForm.height = 495
        $DeletingUsersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $DeletingUsersForm.Controlbox = $false
        $DeletingUsersForm.Icon = $Icon
        $DeletingUsersForm.FormBorderStyle = 'Fixed3D'
        $DeletingUsersForm.Text = "Delete Leavers User Accounts (W10) v$Version."
        $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox1.Location = '40,20'
        $LeaverBox1.size = '650,90'
        $LeaverBox1.text = '1. Check: This deletes Users AD accounts and mailboxes:'
        ### Create group 1 box text labels. ###
        $LeavertextLabel1 = New-Object System.Windows.Forms.Label
        $LeavertextLabel1.Location = '20,25'
        $LeavertextLabel1.size = '600,30'
        $LeavertextLabel1.Text = 'Check the User in the list below to ensure they are marked to be deleted.' 
        $LeavertextLabel2 = New-Object System.Windows.Forms.Label
        $LeavertextLabel2.Location = '20,55'
        $LeavertextLabel2.size = '600,30'
        $LeavertextLabel2.Text = 'IF THE ANY USER IS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
        ### Create group 2 box in form. ###
        $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox2.Location = '40,120'
        $LeaverBox2.size = '650,225'
        $LeaverBox2.text = '2. Check the User below is labelled to be Deleted'
        # Create group 2 box text labels.
        $LeavertextLabel4 = New-Object System.Windows.Forms.Label    
        $LeavertextLabel4.Location = '40,20'
        $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $LeavertextLabel4.size = '570,170'
        $LeavertextLabel4.Text = $LeaversListToDisplayInForm 
        ### Create group 3 box in form. ###
        $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox3.Location = '40,355'
        $LeaverBox3.size = '650,30'
        $LeaverBox3.text = '3. Click Ok to DELETE the User Account or Exit:'
        $LeaverBox3.button
        ### Add an OK button ###
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,395'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '470,395'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        ### Add all the Form controls ### 
        $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
        $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
        #### Activate the form ###
        $DeletingUsersForm.Add_Shown( {$DeletingUsersForm.Activate()})    
        #### Get the results from the button click ###
        $dialogResult = $DeletingUsersForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteLeavers form and starting to Remove User AD Account" -Verbose
            DeleteUserADAccount
        }
    }
    ## create the Select one User AD accounts function complete   ###       
    #
    ## create the Delete User AD account function   ###
    Function DeleteUserADAccount {     
        ## create a list to add the deleted user names to   ###
        $DeleteADUserList = ""
        try {
            ForEach ($UserToDelete in $LeaversOlderThan1MonthList) {
                Write-Verbose "The AD account for user $($UserToDelete.Name) is past its delete date." -Verbose
                Remove-ADObject -Identity $UserToDelete.DistinguishedName -Recursive
                $DelADUserList = $($UserToDelete.Name)
                $DeleteADUserList += "$DelADUserList`r`n"
                Write-Verbose "The AD account for user $($UserToDelete.Name) has been deleted." -Verbose 
                $poplabel = "Deleting the AD account for User`n`n$($UserToDelete.Name)."
                PopupForm
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG deleting the User`n`n$($UserToDelete.Name)`n`nAD account !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Delete User Accounts for Leavers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            ## remove variables ###
            Remove-Variable "UserToDelete", "UserInfo", "DelADUserList", "DeleteADUserList"
            $DeletingUsersForm.Close()
            break 
        }
        ## Details of deleted AD user accounts to include in email to helpdesk   ###
        $DeletedADListForMessage = @()
        $DeletedADListForMessage = $DeleteADUserList
        $DeletedADListForMessageTxt = ''
        $DeletedADListForMessage | ForEach-Object { $DeletedADListForMessageTxt += $_ + "`n" }
        DeleteUserFolders
    }
    ## Delete User AD account function complete   ###
    #
    ## create the Delete Users Folders function   ###
    Function DeleteUserFolders {  
        ## Get User Folders with Delete after Date more than 1 month old   ###
        $DeletedUserList = ""
        $path = "\\scotcourts.local\Home\P"
        $folders = Get-ChildItem $path -Filter "*xx DELETE after DATE*" -directory | select-object Name, PSChildName | select-object PSChildName, @{n = 'Name'; e = {$_.Name -replace '^.*-'}} 
        ForEach ($folder in $folders) {
            $FolderDate = [System.DateTime]$folder.Name
            If ($FolderDate -lt $DateToDel) {
                ## delete user folder on scotcourts.local\Home\P   ###
                try {
                    if (test-path $path\$($folder.PSChildname)) {
                        Remove-Item -path "$path\$($folder.PSChildName)" -Force -Recurse
                        $DelUserList = $($folder.PSChildName) 
                        $DeletedUserList += "$DelUserList`r`n"
                        Write-Verbose "The User folder on SAUFS01 for user $($folder.PSChildname) has been deleted." -Verbose
                        $poplabel = "Deleting the User`n`n$($folder.PSChildname)`n`nP drive folder on $path."
                        PopupForm
                    }
                }       
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG removing the User`n`n$($folder.PSChildname)`n`nP drive folder !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Delete User Accounts for Leavers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    ## remove variables ###
                    Remove-Variable "UserToDelete", "UserInfo", "DelADUserList", "DeleteADUserList"
                    $DeletingUsersForm.Close()
                    break  
                }
            }
        }
        ## Details of deleted user folders to include in email to helpdesk   ###
        $DeletedUserFolderListForMessage = @()
        $DeletedUserFolderListForMessage = $DeletedUserList
        $DeletedUserFolderListForMessageTxt = ''
        $DeletedUserFolderListForMessage | ForEach-Object { $DeletedUserFolderListForMessageTxt += $_ + "`n" }
        DeleteProfileFolders
    }
    ## create the Delete Users Folders function complete   ###
    #
    ## create the Delete Profile Folders function   ###
    Function DeleteProfileFolders {  
        ## Get Profile Folders with Delete after Date more than 1 month old   ###
        $DeletedProfileList = ""
        $path = "\\scotcourts.local\data\profiles"
        $folders = Get-ChildItem $path -Filter "*xx DELETE after DATE*" -directory | select-object Name, PSChildName | select-object PSChildName, @{n = 'Name'; e = {$_.Name -replace '^.*-'}} 
        ForEach ($Profilefolder in $folders) {
            $FolderDate = [System.DateTime]$Profilefolder.Name
            If ($FolderDate -lt $DateToDel) {
                ## delete profile folder on scotcourts.local\data\profiles   ###
                try {
                    if (test-path \\scotcourts.local\data\profiles\$($Profilefolder.PSChildname)) {
                        Remove-Item -path "\\?\$path\$($Profilefolder.PSChildName)" -Force -Recurse
                        $DelProfileList = $($Profilefolder.PSChildName) 
                        $DeletedProfileList += "$DelProfileList`r`n"
                        Write-Verbose "The Profile folder on SAUFS01 for user $($Profilefolder.PSChildname) has been deleted." -Verbose
                        $poplabel = "Deleting the User`n`n$($Profilefolder.PSChildname)`n`nProfile folder on \\scotcourts.local\data\profiles."
                        PopupForm
                    }
                }       
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG removing the User`n`n$($Profilefolder.PSChildname)`n`nProfile folder !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Delete User Accounts for Leavers.',
                        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $DeletingUsersForm.Close()
                    break  
                }
            }
        }
        ## Details of deleted user folders to include in email to helpdesk   ###
        $DeletedProfileFolderListForMessage = @()
        $DeletedProfileFolderListForMessage = $DeletedProfileList
        $DeletedProfileFolderListForMessageTxt = ''
        $DeletedProfileFolderListForMessage | ForEach-Object { $DeletedProfileFolderListForMessageTxt += $_ + "`n" }
        Write-Verbose "The Leavers older than 1 month have been deleted." -Verbose
        ##  send email to helpdesk and infrastructure   ###
        $message2 = "The User P drive folders below have been deleted from SAUFS01.`n$DeletedUserFolderListForMessageTxt`nThe Profile folders below have been deleted from SAUFS01 .`n$DeletedProfileFolderListForMessageTxt`nThe user AD accounts below have been deleted.`n$DeletedADListForMessageTxt"  
        $mailrecipients = "helpdesk@scotcourts.gov.uk", "itinfrastructure@scotcourts.gov.uk"
        Send-MailMessage -To $mailrecipients -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: Leavers - Deleted accounts $(Get-Date -format ('dd MMMM yyyy'))" -Body "$message2" -SmtpServer mail.scotcourts.local
        ##  Message complete message   ###
        [System.Windows.Forms.MessageBox]::Show("The Delete Leavers process is complete.`n`nA list of deleted user accounts will be emailed.", 'Delete leavers.',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        break
    }    
    ## create the Delete Profile Folders function complete   ###
    #
    ## create the main form   ###
    #
    ## Set Delete Date to -1 month   ###
    $DatetoDel = (Get-Date).AddMonths(-1).ToString('dd MMMM yyyy')
    ## get list of users in the Z-Disabled_Leavers OU that are marked to be deleted   ###
    $LeaversList = Get-ADUser -Filter * -SearchBase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties Name, description, DistinguishedName |
        Select-Object Name, SamAccountName, Description, DistinguishedName | Where-Object {$_.Description -match "xx DELETE after DATE"}
    ## get users older than 1 month   ###
    $LeaversOlderThan1MonthList = ForEach ($UserToDelete in $LeaversList) {
        ## get date on users AD description   ###
        $UserInfo = Get-ADUser -identity $UserToDelete.SamAccountName -Properties * |
            Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = {$_.Description -replace '^.*-'}} 
        $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
        ## check if users delete date is over 1 month   ###
        If ($UserDelDate -lt $DateToDel) {
            Get-Aduser $UserToDelete.SamAccountName -Properties * | select-object Name, SamAccountName, Description, DistinguishedName
            Write-Verbose "The AD account for user $($UserToDelete.Name) is past its delete date." -Verbose
        }
    }
    ## if no users to delete show message and exit   ###
    if ($LeaversOlderThan1MonthList -eq $null) {
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
