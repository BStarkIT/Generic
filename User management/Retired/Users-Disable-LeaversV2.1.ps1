# Disable User Accounts for Leavers v2.1
# Author        Brian Stark
# Date          02/09/2020
# Version       2.1
# Purpose       To process the accounts of Leavers.
# Useage        helpdesk staff run a shorcut to Users - Disable Leaver Accounts.
#
# Changes       v2.02 09/12/2019 GK changes to take account of the new personal folder location. 
#                                Added 'ou=SCTS Users,ou=User Accounts,ou=SCTS,DC=scotcourts,DC=local' as another searchbase.     
#               V1.1 11/09/2018 JM oldest refrence of script
#
# Script function: This script performs the following on a User account.
$message = "
AD account - Disables Users account.
AD account - Moves user account to Z-Disabled_Leavers OU.
AD account - Labels account to be deleted in 1 month.
AD account - Removes Profile and Home Folder paths.
Email - Hides email address from Global Address list.
Email - Stops mailbox from receiving email.
Email - Removes Users shared mailbox permissions.
SAUFS01 - Labels user Profile and P drive folders to be deleted in 1 month.
Backup - Gets Users membership of Security Groups and Distribution lists and backs up on users p drive.
Groups - Removes User from all Security groups and Distribution lists."
#
# Start of script:
Start-Transcript -Path "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox\Logs\BS-disable.txt" -append
#
$version='2.1'
#  Create Session with Exchange 2013   #
#$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
#Import-PSSession $session
#
# Get listof UserNames from AD OU's   #
$Users1 = Get-aduser -Filter * -searchbase 'ou=tribunalusers,ou=tribunals,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users2 = Get-aduser -Filter * -searchbase 'ou=sheriffsparttime,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users3 = Get-aduser -Filter * -searchbase 'ou=scs employees,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Where-Object {($_.DistinguishedName -notlike '*OU=deleted users,*') -and ($_.DistinguishedName -notlike '*OU=it administrators,*')} | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users4 = Get-aduser -Filter * -searchbase 'ou=JP,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users5 = Get-aduser -Filter * -searchbase 'ou=judges,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users6 = Get-aduser -Filter * -searchbase 'ou=sheriffs,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname| select-Object -ExpandProperty DisplayName
$Users7 = Get-aduser -Filter * -searchbase 'ou=sheriffsprincipal,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users8 = Get-aduser -Filter * -searchbase 'ou=sheriffssummary,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users9 = Get-aduser -Filter * -searchbase 'ou=sheriffsretired,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users10 = Get-aduser -Filter * -searchbase 'ou=courts,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users11 = Get-aduser -Filter * -searchbase 'ou=judiciary,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users12 = Get-aduser -Filter * -searchbase 'ou=tribunals,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$Users13 = Get-aduser -Filter * -searchbase 'ou=SCTS Users,ou=User Accounts,ou=SCTS,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$UserNameList = $Users1 + $Users2 + $Users3 + $Users4 + $Users5 + $Users6 + $Users7 + $Users8 + $Users9 + $Users10 + $Users11 + $Users12 + $Users13
#  Show Start Message:   #
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$StartMessage = [System.Windows.Forms.MessageBox]::Show("This script performs the following on a User account." + "`n" + "$message", 'Disable User Account for Leavers.',
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
        $PopForm.Text = "Disable User Account for Leavers."
        $PopForm.Size = New-Object System.Drawing.Size(420, 200)
        $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
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
    # create the pop up information form complete   #
    #
    # create the main form   #
    #
    Function LeaverProcess {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        # Set the details of the form. #
        $LeaverProcessForm = New-Object System.Windows.Forms.Form
        $LeaverProcessForm.width = 745
        $LeaverProcessForm.height = 475
        $LeaverProcessForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $LeaverProcessForm.Controlbox = $false
        $LeaverProcessForm.Icon = $Icon
        $LeaverProcessForm.FormBorderStyle = 'Fixed3D'
        $LeaverProcessForm.Text = "Disable User Account for Leavers (W10) v$version."
        $LeaverProcessForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $LeaverBox1 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox1.Location = '40,20'
        $LeaverBox1.size = '650,80'
        $LeaverBox1.text = '1. Select a UserName from the list:'
        $LeavertextLabel1 = New-Object 'System.Windows.Forms.Label';
        $LeavertextLabel1.Location = '20,35'
        $LeavertextLabel1.size = '200,40'
        $LeavertextLabel1.Text = 'Username:' 
        $LeaverMBNameComboBox1 = New-Object System.Windows.Forms.ComboBox
        $LeaverMBNameComboBox1.Location = '275,30'
        $LeaverMBNameComboBox1.Size = '350, 350'
        $LeaverMBNameComboBox1.AutoCompleteMode = 'Suggest'
        $LeaverMBNameComboBox1.AutoCompleteSource = 'ListItems'
        $LeaverMBNameComboBox1.Sorted = $true;
        $LeaverMBNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $LeaverMBNameComboBox1.DataSource = $UserNameList
        $LeaverMBNameComboBox1.Add_SelectedIndexChanged( {$LeaverSelectedMailBoxNametextLabel6.Text = "$($LeaverMBNameComboBox1.SelectedItem.ToString())"})
        $LeaverBox2 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox2.Location = '40,120'
        $LeaverBox2.size = '650,215'
        $LeaverBox2.text = '2. Check the details below are correct before proceeding:'
        $LeavertextLabel3 = New-Object System.Windows.Forms.Label
        $LeavertextLabel3.Location = '20,40'
        $LeavertextLabel3.size = '100,40'
        $LeavertextLabel3.Text = 'The User:'
        $LeavertextLabel4 = New-Object System.Windows.Forms.Label
        $LeavertextLabel4.Location = '330,20'
        $LeavertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $LeavertextLabel4.size = '310,170'
        $LeavertextLabel4.Text = $message
        $LeavertextLabel5 = New-Object System.Windows.Forms.Label
        $LeavertextLabel5.Location = '20,150'
        $LeavertextLabel5.size = '330,40'
        $LeavertextLabel5.Text = 'Will have the following actioned on their account:'
        $LeaverSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
        $LeaverSelectedMailBoxNametextLabel6.Location = '20,80'
        $LeaverSelectedMailBoxNametextLabel6.Size = '350,40'
        $LeaverSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
        $LeaverBox3 = New-Object System.Windows.Forms.GroupBox
        $LeaverBox3.Location = '40,345'
        $LeaverBox3.size = '650,30'
        $LeaverBox3.text = '3. Click Ok to Process User or Exit:'
        $LeaverBox3.button
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,385'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = new-object System.Windows.Forms.Button
        $CancelButton.Location = '470,385'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $LeaverProcessForm.Close()
                $LeaverProcessForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        $LeaverProcessForm.Controls.AddRange(@($LeaverBox1, $LeaverBox2, $LeaverBox3, $OKButton, $CancelButton))
        $LeaverBox1.Controls.AddRange(@($LeavertextLabel1, $LeaverMBNameComboBox1))
        $LeaverBox2.Controls.AddRange(@($LeavertextLabel3, $LeavertextLabel4, $LeavertextLabel5, $LeaverSelectedMailBoxNametextLabel6))
        $LeaverProcessForm.Add_Shown( {$LeaverProcessForm.Activate()})    
        $dialogResult = $LeaverProcessForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            #  CHECK - Don't accept no User selection   # 
            if ($LeaverSelectedMailBoxNametextLabel6.Text -eq '') {
                [System.Windows.Forms.MessageBox]::Show("You need to select a User !!!!!`n`nTrying to enter blank fields is never a good idea.`n`nProcessing cannot continue.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                $LeaverProcessForm.Dispose()
                break
            }
            # CHECK -  User has email address   #
            $MailboxCheck = Get-ADUser -filter {DisplayName -eq $LeaverSelectedMailBoxNametextLabel6.Text} -Properties * | Select-Object EmailAddress
            If ($MailboxCheck.EmailAddress -eq $null) {
                [System.Windows.Forms.MessageBox]::Show("The selected User does not have a mailbox`n`nProcessing cannot continue.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                $LeaverProcessForm.Dispose()
                Write-Verbose "No mailbox back to main form." -Verbose
                Return LeaverProcess 
            }
            #  get primary smtpaddress & sam name from mailbox display name  #
            $MailBoxPrimarySMTPAddress = get-mailbox $($LeaverMBNameComboBox1.SelectedItem.ToString()) | Select-Object primarysmtpaddress| Select-Object -ExpandProperty PrimarySMTPAddress  
            $UserSamAccountName = get-mailbox $($LeaverMBNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName 
            #
            # CHECK - continue if only 1 email address & sam is in pipe if not exit #
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) {LeaverProcessForm}
            if (($UserSamAccountName | Measure-Object).count -ne 1) {LeaverProcessForm}
            #
            # set Date for 1 month ahead  #
            $DateDel = (Get-Date).AddMonths(1).ToString('dd MMM yyyy') 
            #
            # get users membership of Security Groups and Distribution lists &  backup to P drive   #
            $poplabel = "Checking and backing up Users Security group`n`nand`n`nDistribution List membership."
            PopupForm
            $Usermembership = get-aduser -identity $UserSamAccountName -property MemberOf |
                Foreach-Object {$_.MemberOf | Get-AdGroup | Select-Object Name, SamaccountName | Select-object -ExpandProperty SamAccountName}
            $pdrive = test-path \\scotcourts.local\home\P\$UserSamAccountName
            try {
                    If (($null -ne $Usermembership) -and ($pdrive -eq $true)) {                   
                    $Usermembership | out-file \\scotcourts.local\home\P\$UserSamAccountName\UserMembershipBackup.csv
                    Write-Verbose "Backed up Users Security group and Distribution List membership to users P drive." -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG backing up the users group membership !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess  
            }
            #
            # change profile and P drive folders name - add delete in 1 month to folder names  #
            $poplabel = "Renaming Users Profile folder and P drive folders`n`non SAUFS01`n`nwith Delete in 1 month."
            PopupForm
            try {
                if (test-path \\scotcourts.local\data\profiles\"$UserSamAccountName.v2") {
                    Get-Item \\scotcourts.local\data\profiles\"$UserSamAccountName.v2" | Rename-Item -NewName {$_.Name -replace "$", " xx DELETE after DATE - $Datedel"}
                    Write-Verbose "Renamed Users Profile folder on SAUFS01" -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG with renaming the users profile folder !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Processing User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess
            }
            try {
                if (test-path \\scotcourts.local\Home\P\$UserSamAccountName) {
                    Get-Item \\scotcourts.local\home\P\$UserSamAccountName | Rename-Item -NewName {$_.Name -replace "$", " xx DELETE after DATE - $Datedel"}
                    Write-Verbose "Renamed users P drive folder on SAUFS01" -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG with renaming the users P drive !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Processing User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess
            }
            #
            # Hide email address from Address list & accept only from helpdesk #
            $poplabel = "Hiding email address from global address list`n`nand`n`nSetting users mailbox to not receive emails"
            PopupForm
            try {
                Get-Mailbox $MailBoxPrimarySMTPAddress | Set-Mailbox -AcceptMessagesOnlyFrom 'helpdesk' -HiddenFromAddressListsEnabled $true
                Write-Verbose "Setting Users mailbox not to receive emails and hiding from address list" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG setting the users mailbox !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Processing User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess  
            }
            #
            # disable user AD account   #
            $poplabel = "Disabling Users account."
            PopupForm
            try {
                Disable-ADAccount -Identity $UserSamAccountName
                Write-Verbose "Disabling AD Account" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG disabling the users AD account !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Processing User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess 
            }
            #
            # edit AD description field to delete account in 1 month & clear profile & P drive paths   #
            $poplabel = "Editing users AD account:`n`nLabelling to be deleted in 1 month.`n`nClearing Profile Path and P drive path."
            PopupForm
            try {
                Get-ADUser $UserSamAccountName -Properties Description | ForEach-Object {Set-ADUser $_ -Description "$($_.Description) xx DELETE after DATE - $Datedel"}
                Get-ADUser $UserSamAccountName -Properties ProfilePath | ForEach-Object {Set-ADUser $_ -ProfilePath $Null}
                Get-ADUser $UserSamAccountName -Properties HomeDirectory | ForEach-Object {Set-ADUser $_ -HomeDirectory $Null}
                Get-ADUser $UserSamAccountName -Properties HomeDrive | ForEach-Object {Set-ADUser $_ -HomeDrive $Null} 
                Write-Verbose " Labelling users AD account to be deleted in 1 month and clearing Profile and P drive paths" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG labelling the users AD account to be deleted in 1 month !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess
            }
            #
            # move ad account to Z-Disabled Leavers ou   #
            $poplabel = "Moving user account to the `n`nSCTS/User Accounts/Z-Disabled_Leavers OU."
            PopupForm
            try {
                Get-ADUser $UserSamAccountName | Move-ADObject -targetpath 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
                Write-Verbose "moving Users AD account to Z-Disabled_Leavers OU" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG moving the users account to the Z-Disabled_Leavers OU !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess 
            }
            #
            # get user shared mailbox access   #
            $poplabel = "Checking and removing`n`nany Users Shared Mailbox permissions"
            PopupForm
            $SharedMailboxes = Get-ADUser -Identity $UserSamAccountName -Properties msExchDelegateListBL | Select-Object -ExpandProperty msExchDelegateListBL
            $SharedMailboxes -replace '^CN=|,.*$'
            $SharedMailboxList = $SharedMailboxes | ForEach-Object {get-mailbox -identity $_ | select-object PrimarySMTPAddress}
            # Remove shared mailbox Full access for user   #
            try {
                If ($SharedMailboxList -ne $null) {
                    $SharedMailboxList | ForEach-Object {                    
                        Remove-MailboxPermission $_.PrimarySMTPAddress -User $UserSamAccountName -AccessRights FullAccess -confirm:$false 
                        Set-Mailbox $_.PrimarySMTPAddress -GrantSendOnBehalfTo @{remove = "$UserSamAccountName"}
                        Write-Verbose "Removed Users Shared Mailbox permissions" -Verbose
                    }
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG removing the users shared mailbox permissions !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess
            }
            #
            # Remove User from Security groups and Distribution lists   #
            $poplabel = "Checking and removing User from`n`nSecurity groups`n`nand Distribution Lists."
            PopupForm
            try {
                If ($null -ne $Usermembership) {                    
                    $Usermembership | ForEach-Object {
                        Remove-ADGroupMember -Identity $_ -Member $UserSamAccountName -confirm:$false
                        Write-Verbose "Removed user from groups" -Verbose}
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG removing the users from the groups !!!.`n`nPlease contact the Systems Integration Team with the details.", 'Disable User Accounts for Leavers.',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $LeaverProcessForm.Close()
                Return LeaverProcess
            }
            [System.Windows.Forms.MessageBox]::Show("Disable User Accounts for Leavers for $($LeaverMBNameComboBox1.SelectedItem.ToString()) - Completed.", 'Disable User Accounts for Leavers.',
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $LeaverProcessForm.Close()
            Return LeaverProcess
            # Leaver processing completed   # 
        }   
    }
    LeaverProcess
}
