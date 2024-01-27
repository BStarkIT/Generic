# Delete Disabled Admin Accounts
# Author        Brian Stark
# Date          10/06/2021
# Version       1.0
# Purpose       To remove Leavers AD user accounts & Profile & Users folders on SAUFS01.
# Useage        helpdesk staff run a shorcut to Removing User Accounts for Leavers.exe
#
# Changes       10/06/2021  BS  1.0 Admin Fork from 2.1 delete user
#               02/09/2020  BS  2.1 Update & clean up of script & fix issues.
#               09/01/2020  GK  1.2 Changes to take acount of W10 OU users. Improves readability.
#               05/12/2019  GK  2.03 Replace references to \\scotcourts.local\data\users with \\scotcourts.local\home\p. Replace references to changed string with variable 
#               06/12/2018  JM  1. user accounts with blackberries won't delete with "remove-aduser" changed to "remove-adobject" & added DeleteUserADAccount function.
#                               2. some user profile folders have filenames longer than 255 characters changed Remove-Item -path to "\\?\$path" in DeleteProfileFolders function.    
#
# Script function: This script performs the following on a User account.
$message = "

This script searches the Z-Disabled_Leavers OU in AD for Admin accounts marked for deletion.

If the 'delete after date' has passed the admin account will be:

AD: Users AD account will be deleted."

#
# Start of script:
#
$version='1.0'
$WinTitle = "Delete Disabled Admin Accounts v$version."
#  Show Start Message:   #
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", $WinTitle, 
    [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {
    Exit
} 
else {
    # create the pop up information form   #
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        # create form
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = $WinTitle
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
    # create the pop up information form complete   #
    #
    # create the Select Multiple User AD account function   #
    Function DeleteMultipleAdminAccount {
        # get list of users with name and description to display in form   #
        $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) {$($User.Name) + " - " + $($User.Description)}
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
        $DeletingUsersForm.Text = $WinTitle
        $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox1.Location = '40,20'
        $LeaverBox1.size = '650,90'
        $LeaverBox1.text = '1. Check: This deletes Admin AD accounts:'
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
        $LeaverBox2.text = '2. Check the Admins below are labelled to be Deleted'
        $LeavertextLabel4 = New-Object System.Windows.Forms.ListBox    
        $LeavertextLabel4.Location = '40,20'
        $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $LeavertextLabel4.size = '570,170'
        $LeavertextLabel4.Datasource = $LeaversListToDisplayInForm 
        $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox3.Location = '40,355'
        $LeaverBox3.size = '650,30'
        $LeaverBox3.text = '3. Click Ok to DELETE Admin Accounts or Exit:'
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
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
        $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
        $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
        $DeletingUsersForm.Add_Shown( {$DeletingUsersForm.Activate()})    
        $dialogResult = $DeletingUsersForm.ShowDialog()
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteLeavers form and starting to Remove User AD Account" -Verbose
            DeleteUserADAccount
        }
    }
    # create the Select Multiple User AD accounts function complete
    #
    # create the Select One User AD account function   #
    Function DeleteOneAdminAccount {
        # get users with name and description to display in form   #
        $LeaversListToDisplayInForm = foreach ($user in $LeaversOlderThan1MonthList) {$($User.Name) + " - " + $($User.Description)}
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $DeletingUsersForm = New-Object System.Windows.Forms.Form
        $DeletingUsersForm.width = 745
        $DeletingUsersForm.height = 495
        $DeletingUsersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $DeletingUsersForm.Controlbox = $false
        $DeletingUsersForm.Icon = $Icon
        $DeletingUsersForm.FormBorderStyle = 'Fixed3D'
        $DeletingUsersForm.Text = $WinTitle
        $DeletingUsersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox1.Location = '40,20'
        $LeaverBox1.size = '650,90'
        $LeaverBox1.text = '1. Check: This deletes Admin AD accounts:'
        $LeavertextLabel1 = New-Object System.Windows.Forms.Label
        $LeavertextLabel1.Location = '20,25'
        $LeavertextLabel1.size = '600,30'
        $LeavertextLabel1.Text = 'Check the Admin in the list below to ensure they are marked to be deleted.' 
        $LeavertextLabel2 = New-Object System.Windows.Forms.Label
        $LeavertextLabel2.Location = '20,55'
        $LeavertextLabel2.size = '600,30'
        $LeavertextLabel2.Text = 'IF THE ANY USER IS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
        $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox2.Location = '40,120'
        $LeaverBox2.size = '650,225'
        $LeaverBox2.text = '2. Check the Admin below is labelled to be Deleted'
        $LeavertextLabel4 = New-Object System.Windows.Forms.Label    
        $LeavertextLabel4.Location = '40,20'
        $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $LeavertextLabel4.size = '570,170'
        $LeavertextLabel4.Text = $LeaversListToDisplayInForm 
        $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox3.Location = '40,355'
        $LeaverBox3.size = '650,30'
        $LeaverBox3.text = '3. Click Ok to DELETE the Admin Account or Exit:'
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
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        $DeletingUsersForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
        $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeavertextLabel2))
        $LeaverBox2.Controls.AddRange(@($LeavertextLabel4))
        $DeletingUsersForm.Add_Shown( {$DeletingUsersForm.Activate()})    
        $dialogResult = $DeletingUsersForm.ShowDialog()
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteLeavers form and starting to Remove User AD Account" -Verbose
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
                Write-Verbose "The AD account for user $($UserToDelete.Name) is past its delete date." -Verbose
                Remove-ADUser -Identity $UserToDelete.SamAccountName
                # original remove aduser to be removed  #
                $DelADUserList = $($UserToDelete.Name)
                $DeleteADUserList += "$DelADUserList`r`n"
                Write-Verbose "The AD account for user $($UserToDelete.Name) has been deleted." -Verbose 
                $poplabel = "Deleting the AD account for User`n`n$($UserToDelete.Name)."
                PopupForm
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG deleting the User`n`n$($UserToDelete.Name)`n`nAD account !!!.`n`nPlease contact the Systems Integration Team with the details.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            # remove variables #
            #Remove-Variable "UserToDelete", "UserInfo", "DelADUserList", "DeleteADUserList"
            $DeletingUsersForm.Close()
            break 
        }
        Return
    }
    # Set Delete Date to -1 month   #
    $DatetoDel = (Get-Date).ToString('dd MMMM yyyy')
    # get list of users in the Z-Disabled_Leavers OU that are marked to be deleted   #
    $LeaversList = Get-ADUser -Filter * -SearchBase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties Name, description, DistinguishedName |
        Select-Object Name, SamAccountName, Description, DistinguishedName | Where-Object {$_.Description -match "YY DELETE after DATE"}
    # get users older than 1 month   #
    $LeaversOlderThan1MonthList = ForEach ($UserToDelete in $LeaversList) {
        # get date on users AD description   #
        $UserInfo = Get-ADUser -identity $UserToDelete.SamAccountName -Properties * |
            Select-Object Name, SamAccountName, Description, DistinguishedName | select-object SamAccountName, @{n = 'DeleteUserOnDate'; e = {$_.Description -replace '^.*-'}} 
        $UserDelDate = [System.DateTime]$UserInfo.DeleteUserOnDate
        # check if users delete date is over 1 month   #
        If ($UserDelDate -lt $DateToDel) {
            Get-Aduser $UserToDelete.SamAccountName -Properties * | select-object Name, SamAccountName, Description, DistinguishedName
            Write-Verbose "The Admin account for user $($UserToDelete.Name) is past its delete date." -Verbose
        }
    }
    # if no users to delete show message and exit   #
    if ($LeaversOlderThan1MonthList -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("There are currently no Admin to delete.`n`nThe Delete Admin process is complete.", $WinTitle,
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        break
    }
    if (@($LeaversOlderThan1MonthList).Count -eq 1) {
        DeleteOneAdminAccount
    }
    Else {
        DeleteMultipleAdminAccount
    }
}
