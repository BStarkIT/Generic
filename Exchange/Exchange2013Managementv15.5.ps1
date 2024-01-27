# Exchange 2013 Management v1.55
#region Info
# Author        Brian Stark
# Date          14/04/2020
# Version       1.55
# Purpose       To manage exchange 2013 shared mailboxes & calendars.
# Useage        helpdesk staff run a shorcut to Exchange 2013 Management.exe
#               Script loads exchange 2013 session & imports users & mailboxes & distribution lists from AD.
# Revisions     
#               V1.55 19/06/2020 BS - regions added & corrected
#               v1.54 18/06/2020 BS - Additon of enabling a copy of sent emil to sent folder of shared mailbox.
#               v1.53 05/05/2020 BS - addition of seconday location for shared mailboxes
#               V1.52 14/04/2020 BS - addition of ActiveSync enable & device removal needed for Blackberry enablement & rebuild
#               V1.51 06/04/2020 BS - updated registry paths for Win10 & office 2016, corrected user/Admin user SID error 
#               v1.5 30/03/2020 BS - updated to office 2016, added Win10 user groups, SCTS Distribution lists, revised commands for Powershell v5.1 syntax & replaced commands blocked by GPO
#               v1.44 03/09/2018 JM - can't add send on behalf permissions for more than 1 user on new shared mailboxes. Can add more than 1 on existing mailboxes. Edited match in line 414 to If ($Status -match "$UserDisplayName.*")
#               v1.43 02/08/2018 JM - can't add send on behalf permissions for more than 1 user. If you try to add 2 users message that 2nd user already has access.Edited match in line 412 to If (($Status -match "$UserDisplayName.*").count -eq "1").
#               v1.42 20/07/2018 JM - "check shared mailbox send on behalf of" is displaying aberdeen users.the aberdeen mailbox was used for testing & has been left in script. changed mailbox to $mailbox.Name.
#               v1.41 18/07/2018 JM - When script opened in Visual Studio script won't run.Replaced "s with 's in script to use script in Visual Studio Code.
#               Shared Mailbox management added 1.7 Shared Mailbox LogOn line 974 & removed "6.Shared Mailbox - Autoreply (Outlook Rule not Out of Office)" and "7.Shared mailbox - Out of Office turn off and on".
#               Updated "check" functions to remove scs\ & scotcourts.local & just show user names.
#               v1.40 01/06/2018 - Distribution Lists add & remove users added '-SearchBase OU=Distribution Lists,OU=Groups,OU=COURTS,DC=scotcourts,DC=local' to specify OU in AD in lines 2778 & 2940.
#               v1.39 04/05/2018 OU's in AD have been moved & renamed. Updated 'New shared mailbox & new trib shared mailbox'
#               v1.38 01/05/2018 OU's in AD have been moved & renamed. Updated 'Get Shared mailboxes & Distribution Lists & Users from AD'. 
#               v1.37 18/04/2018 corrected typo in line 799 should be SharedMailboxManagementForm
#               v1.36 16/04/2018 helpdesk highlighted in 'Disabled User Account form' if selected user has no email address code continues to run & sets 2 options for multiple ad users.
#               added check user has email address line 3995 to stop if user has no email address.
#               added 'return' where functions are called to stop script continuing.
#               v1.35 13/02/2018 JM changed send email to helpdesk to use $env:UserName instead of helpdesk.  
#               v1.34 06/12/2017 JM Change 1:1.7 Add New shared mailbox changed to 1.7 Add New Courts Shared Mailbox 1.8 Add New Tribunals shared mailbox. 
#               Change 2: User Out of Office SubForm - Add message and turn On changed box size & edited scrollbar setting.
#               v1.33 06/11/2017 JM added $users10,11 & 12 line 98 to pick up users in the new Courts OU.
#               v1.32 02/11/17 JM typo line 939 Set-Mailbox $MailBoxPrimarySMTPAddress@scotcourts.gov.uk changed to Set-Mailbox $MailBoxPrimarySMTPAddress 
#               v1.31 New Shared mailbox added Set-ADUser to add Owner name & description.
#       
#               v1.3 24/10/2017 JM
#               Change1: Added 'Add New Distribution List' form to 'Distribution list management'.
#               Change2: Added 'Add New Shared Mailbox' form to 'Shared Mailbox management'.
#               Change3: Changed 'Don't accept null username or mailbox' added 'break' to exit. If user clicked ok with null field then clicked cancel script would continue to run.
#               Change4: Changed Distribution list Add,Remove & List to use samaccountname.
#       
#               v1.2 28 28/08/17 JM
#               Change 1:Send email to helpdesk added to '1.5 Remove- Full Access permissions for a User' and '1.6 Remove - Send On Behalf Of permissions for a User'
#               Change 2: -Wait option added to Out-GridView to get out gridview to front of screen to '1.3 Check - Mailbox current Full Access permissions','1.4 Check - Mailbox current Send On Behalf Of permissions',
#                         '3.3 Check - current calendar permissions',' 4.3 Distribution List - List current members of a List' 
#               Change 3: 'Shared calendar management' form and subforms added.
#               Change 4: 'Shared mailbox Out of Office' form and subforms added.
#
#               v1.1 23/08/17 JM (Line numbers are in Exchange 2013 Management v1.0) 
#               3 changes to 'Disabled User Account form' Line 2208.
#               Change 1: Typo in line 2308 '$HideAcceptMBNameComboBox2' changed to '$HideAcceptUserNameComboBox1'.
#               Change 2: Line 2317 'Message Complete message' text changed to 'The mailbox has been hidden from the Address Lists.It is now only accepting incoming email from the IT helpdesk'    
#               Change 3: Line 2321 Added line to add date to include date in email sent to helpdesk.   
#
#               10 Changes to add a check that only 1 mailbox is in pipe before actions applied 'if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) {Mainform}'.
#               Change 4: Line 202 - Add full mailbox permission form 
#               Change 5: Line 359 - Add sendonbehalfof mailbox permission form
#               Change 6: Line 725 - remove full access permission
#               Change 7: Line 877 - Remove sendonbehalfof permission
#               Change 8: Line 1528 - Out of office turn on
#               Change 9: Line 1644 - Out of office turn off
#               Change 10: Line 1768 - Out of Office add message and turn on
#               Change 11: Line 1034 - Distribution list add user 
#               Change 12: Line 1183 - Distribution list remove user  
#               Change 13: Line 2308 - Disable User remove from address list 
#
#endregion Info
#region Variables
#####################################################################
########    Create Session with Exchange 2013    ####################
#####################################################################
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
########################################################################################################
################           Set icon for all forms and subforms         #################################
########################################################################################################
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\email.ico'
########################################################################################################
################     Get Shared mailboxes & Distribution Lists from AD    ##############################
########################################################################################################
$Sharedmailbox1 = Get-Mailbox -OrganizationalUnit 'ou=shared mailboxes,ou=resource accounts,ou=useraccounts,ou=courts,dc=scotcourts,dc=local' | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName 
$Sharedmailbox2 = Get-Mailbox -OrganizationalUnit 'ou=user accounts (shared),ou=scts,dc=scotcourts,dc=local' | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName 
$Sharedmailboxes = $Sharedmailbox2 + $Sharedmailbox1
$Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name  
#########################################################################################################
#################             Get listof UserNames from AD OU's            ##############################
#########################################################################################################
$Users1 = Get-ADUser –Filter * -SearchBase 'ou=tribunalusers,ou=tribunals,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users2 = Get-ADUser –Filter * -SearchBase 'ou=sheriffsparttime,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users3 = Get-ADUser –Filter * -SearchBase 'ou=scs employees,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Where-Object { ($_.DistinguishedName -notlike '*OU=deleted users,*') -and ($_.DistinguishedName -notlike '*OU=it administrators,*') } | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users4 = Get-ADUser –Filter * -SearchBase 'ou=JP,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users5 = Get-ADUser –Filter * -SearchBase 'ou=judges,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users6 = Get-ADUser –Filter * -SearchBase 'ou=sheriffs,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users7 = Get-ADUser –Filter * -SearchBase 'ou=sheriffsprincipal,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users8 = Get-ADUser –Filter * -SearchBase 'ou=sheriffssummary,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users9 = Get-ADUser –Filter * -SearchBase 'ou=sheriffsretired,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users10 = Get-ADUser –Filter * -SearchBase 'ou=courts,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users11 = Get-ADUser –Filter * -SearchBase 'ou=judiciary,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users12 = Get-ADUser –Filter * -SearchBase 'ou=tribunals,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users13 = Get-ADUser –Filter * -SearchBase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$UserNameList = $Users1 + $users2 + $users3 + $users4 + $users5 + $users6 + $users7 + $users8 + $users9 + $users10 + $users11 + $users12 + $users13
#########################################################################################################
#
#endregion Variables
#region Subforms
#########################################################################################################
###############       Create 'Shared Mailbox - User Access Management' Sub Forms        #################
#########################################################################################################
#
######################################################################
####  Create SubForm 'Add - Full Access permissions for a User'.   ###
######################################################################
#region Fullaccess
Function AddMBFullAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddMBFullForm = New-Object System.Windows.Forms.Form
    $AddMBFullForm.width = 780
    $AddMBFullForm.height = 500
    $AddMBFullForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddMBFullForm.Controlbox = $false
    $AddMBFullForm.Icon = $Icon
    $AddMBFullForm.FormBorderStyle = 'Fixed3D'
    $AddMBFullForm.Text = 'Mailbox - Add Full Access Permissions for a User.'
    $AddMBFullForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddMBFullBox1 = New-Object System.Windows.Forms.GroupBox
    $AddMBFullBox1.Location = '40,40'
    $AddMBFullBox1.size = '700,125'
    $AddMBFullBox1.text = '1. Select a UserName and MailBoxName from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $AddMBFulltextLabel1 = New-Object System.Windows.Forms.Label
    $AddMBFulltextLabel1.Location = '20,40'
    $AddMBFulltextLabel1.size = '150,40'
    $AddMBFulltextLabel1.Text = 'UserName:' 
    $AddMBFulltextLabel2 = New-Object System.Windows.Forms.Label
    $AddMBFulltextLabel2.Location = '20,80'
    $AddMBFulltextLabel2.size = '150,40'
    $AddMBFulltextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $AddMBFullUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $AddMBFullUserNameComboBox1.Location = '325,35'
    $AddMBFullUserNameComboBox1.Size = '350, 310'
    $AddMBFullUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $AddMBFullUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $AddMBFullUserNameComboBox1.Sorted = $true;
    $AddMBFullUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddMBFullUserNameComboBox1.DataSource = $UsernameList
    $AddMBFullUserNameComboBox1.add_SelectedIndexChanged( { $AddMBFullSelectedUserNametextLabel4.Text = "$($AddMBFullUserNameComboBox1.SelectedItem.ToString())" })
    $AddMBFullMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddMBFullMBNameComboBox2.Location = '325,75'
    $AddMBFullMBNameComboBox2.Size = '350, 350'
    $AddMBFullMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddMBFullMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddMBFullMBNameComboBox2.Sorted = $true;
    $AddMBFullMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddMBFullMBNameComboBox2.DataSource = $Sharedmailboxes 
    $AddMBFullMBNameComboBox2.add_SelectedIndexChanged( { $AddMBFullSelectedMailBoxNametextLabel6.Text = "$($AddMBFullMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddMBFullBox2 = New-Object System.Windows.Forms.GroupBox
    $AddMBFullBox2.Location = '40,190'
    $AddMBFullBox2.size = '700,125'
    $AddMBFullBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $AddMBFulltextLabel3 = New-Object System.Windows.Forms.Label
    $AddMBFulltextLabel3.Location = '40,40'
    $AddMBFulltextLabel3.size = '100,40'
    $AddMBFulltextLabel3.Text = 'The User:' 
    $AddMBFullSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $AddMBFullSelectedUserNametextLabel4.Location = '30,80'
    $AddMBFullSelectedUserNametextLabel4.Size = '200,40'
    $AddMBFullSelectedUserNametextLabel4.ForeColor = 'Blue'
    $AddMBFulltextLabel5 = New-Object System.Windows.Forms.Label
    $AddMBFulltextLabel5.Location = '275,40'
    $AddMBFulltextLabel5.size = '400,40'
    $AddMBFulltextLabel5.Text = 'Will have Full Access permissions added to the mailbox:'
    $AddMBFullSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $AddMBFullSelectedMailBoxNametextLabel6.Location = '350,80'
    $AddMBFullSelectedMailBoxNametextLabel6.Size = '200,40'
    $AddMBFullSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $AddMBFullBox3 = New-Object System.Windows.Forms.GroupBox
    $AddMBFullBox3.Location = '40,340'
    $AddMBFullBox3.size = '700,30'
    $AddMBFullBox3.text = '3. Click Ok to add Mailbox Full Access permissions or Cancel:'
    $AddMBFullBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $AddMBFullForm.Close()
            $AddMBFullForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $AddMBFullForm.Controls.AddRange(@($AddMBFullBox1, $AddMBFullBox2, $AddMBFullBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $AddMBFullBox1.Controls.AddRange(@($AddMBFulltextLabel1, $AddMBFulltextLabel2, $AddMBFullUserNameComboBox1, $AddMBFullMBNameComboBox2))
    $AddMBFullBox2.Controls.AddRange(@($AddMBFulltextLabel3, $AddMBFullSelectedUserNametextLabel4, $AddMBFulltextLabel5, $AddMBFullSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $AddMBFullForm.AcceptButton = $OKButton
    $AddMBFullForm.CancelButton = $CancelButton
    #### Activate the form ###
    $AddMBFullForm.Add_Shown( { $AddMBFullForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $AddMBFullForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($AddMBFullSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Add Full Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBFullForm.Close()
            $AddMBFullForm.Dispose()
            break
        }
        Elseif ($AddMBFullSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Shared mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Add Full Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBFullForm.Close()
            $AddMBFullForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($AddMBFullUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailBoxPrimarySMTPAddress = get-mailbox $($AddMBFullMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress
        #####################################################################
        #########        Check if User already has full access  #############
        #####################################################################
        $Status = Get-Mailbox $MailBoxPrimarySMTPAddress | Get-MailboxPermission -User $UserSamAccountName
        If ($Status.AccessRights -eq 'FullAccess') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddMBFullUserNameComboBox1.SelectedItem.ToString()) ) Already has Full Access Permissions to the ( $($AddMBFullMBNameComboBox2.SelectedItem.ToString()) ) mailbox.", 'Mailbox - Add Full Access Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBFullForm.Close()
            $AddMBFullForm.Dispose()
            Return SharedMailboxManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######                Add Full access for user        ###############
            #####################################################################
            Add-MailboxPermission –Identity $MailBoxPrimarySMTPAddress -User $UserSamAccountName -AccessRights FullAccess -confirm:$false
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddMBFullUserNameComboBox1.SelectedItem.ToString()) )`nhas had Full Access Permissions added to the ( $($AddMBFullMBNameComboBox2.SelectedItem.ToString()) ) mailbox. `nThe mailbox will appear in a few minutes.", 'Mailbox - Add Full Access Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($AddMBFullUserNameComboBox1.SelectedItem.ToString()) has had Full Access permissions added to the $($AddMBFullMBNameComboBox2.SelectedItem.ToString()) mailbox" -Body 'The user needs to close and re-open Outlook and  the mailbox will appear in a few minutes.' -SmtpServer mail.scotcourts.local
            $AddMBFullForm.Close()
            $AddMBFullForm.Dispose()
            Return SharedMailboxManagementForm
        }
    }
}
#########################################################################
##  Completed -Create SubForm Add - Full Access permissions for a User ##
#########################################################################
#endregion Fullaccess
#
#########################################################################
#####   Create SubForm Add Mailbox Send On Behalf Of permission    ######
#########################################################################
#region SendOB
Function AddMBSendBehalfOfAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddMBSendBehalfOfForm = New-Object System.Windows.Forms.Form
    $AddMBSendBehalfOfForm.width = 745
    $AddMBSendBehalfOfForm.height = 475
    $AddMBSendBehalfOfForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddMBSendBehalfOfForm.Controlbox = $false
    $AddMBSendBehalfOfForm.Icon = $Icon
    $AddMBSendBehalfOfForm.FormBorderStyle = 'Fixed3D'
    $AddMBSendBehalfOfForm.Text = 'Mailbox - Add Send On Behalf Of Access Permissions for a User.'
    $AddMBSendBehalfOfForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddMBSendBehalfOfBox1 = New-Object System.Windows.Forms.GroupBox
    $AddMBSendBehalfOfBox1.Location = '40,20'
    $AddMBSendBehalfOfBox1.size = '650,125'
    $AddMBSendBehalfOfBox1.text = '1. Select a UserName and MailBoxName from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $AddMBSendBehalfOftextLabel1 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOftextLabel1.Location = '20,40'
    $AddMBSendBehalfOftextLabel1.size = '100,40'
    $AddMBSendBehalfOftextLabel1.Text = 'UserName:' 
    $AddMBSendBehalfOftextLabel2 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOftextLabel2.Location = '20,80'
    $AddMBSendBehalfOftextLabel2.size = '100,40'
    $AddMBSendBehalfOftextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $AddMBSendBehalfOfUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $AddMBSendBehalfOfUserNameComboBox1.Location = '275,35'
    $AddMBSendBehalfOfUserNameComboBox1.Size = '350, 310'
    $AddMBSendBehalfOfUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $AddMBSendBehalfOfUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $AddMBSendBehalfOfUserNameComboBox1.Sorted = $true;
    $AddMBSendBehalfOfUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddMBSendBehalfOfUserNameComboBox1.DataSource = $UsernameList
    $AddMBSendBehalfOfUserNameComboBox1.add_SelectedIndexChanged( { $AddMBSendBehalfOfSelectedUserNametextLabel4.Text = "$($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString())" })
    $AddMBSendBehalfOfMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddMBSendBehalfOfMBNameComboBox2.Location = '275,75'
    $AddMBSendBehalfOfMBNameComboBox2.Size = '350, 350'
    $AddMBSendBehalfOfMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddMBSendBehalfOfMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddMBSendBehalfOfMBNameComboBox2.Sorted = $true;
    $AddMBSendBehalfOfMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddMBSendBehalfOfMBNameComboBox2.DataSource = $Sharedmailboxes 
    $AddMBSendBehalfOfMBNameComboBox2.add_SelectedIndexChanged( { $AddMBSendBehalfOfSelectedMailBoxNametextLabel6.Text = "$($AddMBSendBehalfOfMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddMBSendBehalfOfBox2 = New-Object System.Windows.Forms.GroupBox
    $AddMBSendBehalfOfBox2.Location = '40,170'
    $AddMBSendBehalfOfBox2.size = '650,125'
    $AddMBSendBehalfOfBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $AddMBSendBehalfOftextLabel3 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOftextLabel3.Location = '40,40'
    $AddMBSendBehalfOftextLabel3.size = '100,40'
    $AddMBSendBehalfOftextLabel3.Text = 'The User:' 
    $AddMBSendBehalfOfSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOfSelectedUserNametextLabel4.Location = '30,80'
    $AddMBSendBehalfOfSelectedUserNametextLabel4.Size = '200,40'
    $AddMBSendBehalfOfSelectedUserNametextLabel4.ForeColor = 'Blue'
    $AddMBSendBehalfOftextLabel5 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOftextLabel5.Location = '275,40'
    $AddMBSendBehalfOftextLabel5.size = '370,40'
    $AddMBSendBehalfOftextLabel5.Text = 'Will have Send On Behalf Of Access permissions added to:'
    $AddMBSendBehalfOfSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $AddMBSendBehalfOfSelectedMailBoxNametextLabel6.Location = '350,80'
    $AddMBSendBehalfOfSelectedMailBoxNametextLabel6.Size = '200,40'
    $AddMBSendBehalfOfSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $AddMBSendBehalfOfBox3 = New-Object System.Windows.Forms.GroupBox
    $AddMBSendBehalfOfBox3.Location = '40,320'
    $AddMBSendBehalfOfBox3.size = '650,30'
    $AddMBSendBehalfOfBox3.text = '3. Click Ok to add Mailbox Send On Behalf Of Access permissions or Cancel:'
    $AddMBSendBehalfOfBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $AddMBSendBehalfOfForm.Close()
            $AddMBSendBehalfOfForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $AddMBSendBehalfOfForm.Controls.AddRange(@($AddMBSendBehalfOfBox1, $AddMBSendBehalfOfBox2, $AddMBSendBehalfOfBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $AddMBSendBehalfOfBox1.Controls.AddRange(@($AddMBSendBehalfOftextLabel1, $AddMBSendBehalfOftextLabel2, $AddMBSendBehalfOfUserNameComboBox1, $AddMBSendBehalfOfMBNameComboBox2))
    $AddMBSendBehalfOfBox2.Controls.AddRange(@($AddMBSendBehalfOftextLabel3, $AddMBSendBehalfOfSelectedUserNametextLabel4, $AddMBSendBehalfOftextLabel5, $AddMBSendBehalfOfSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $AddMBSendBehalfOfForm.AcceptButton = $OKButton
    $AddMBSendBehalfOfForm.CancelButton = $CancelButton
    #### Activate the form ###
    $AddMBSendBehalfOfForm.Add_Shown( { $AddMBSendBehalfOfForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $AddMBSendBehalfOfForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($AddMBSendBehalfOfSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Add Send On Behalf Of Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBSendBehalfOfForm.Close()
            $AddMBSendBehalfOfForm.Dispose()
            break
        }
        Elseif ($AddMBSendBehalfOfSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Shared mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Add Send On Behalf Of Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBSendBehalfOfForm.Close()
            $AddMBSendBehalfOfForm.Dispose()
            break
        }
        #####################################################################
        ########  get user sama & displayname from user name:  ############## 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $UserDisplayName = get-mailbox $($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString()) | Select-Object Name | Select-Object -ExpandProperty Name
        $MailBoxPrimarySMTPAddress = get-mailbox $($AddMBSendBehalfOfMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        #########   Check if User already has SendOnBehalf access  ##########
        #####################################################################
        $Status = Get-Mailbox $MailBoxPrimarySMTPAddress | Select-Object -ExpandProperty GrantSendOnBehalfTo
        If ($Status -match "$UserDisplayName.*") {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString()) ) already has SendOnBehalfOf permissions to the ( $($AddMBSendBehalfOfMBNameComboBox2.SelectedItem.ToString()) ) mailbox.", 'Mailbox - Add SendOnbehalfOf Permission for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddMBSendBehalfOfForm.Close()
            $AddMBSendBehalfOfForm.Dispose()
            Return SharedMailboxManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######         Add SendOnBehalfOf access for user     ###############
            #####################################################################
            Set-Mailbox $MailBoxPrimarySMTPAddress -GrantSendOnBehalfTo @{add = "$UserSamAccountName" }
            #
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString()) ) has had SendOnBehalfof Permissions added to the ( $($AddMBSendBehalfOfMBNameComboBox2.SelectedItem.ToString()) ) mailbox.", 'Mailbox - Add SendOnbehalfOf Permission for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #########################################################################################################
            #############     Send email to helpdesk to remove old email addresses      #############################
            #########################################################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($AddMBSendBehalfOfUserNameComboBox1.SelectedItem.ToString()) has had Send On Behalf Of permissions added to the $($AddMBSendBehalfOfMBNameComboBox2.SelectedItem.ToString()) mailbox" -Body 'The user needs to close and re-open Outlook and  the mailbox will appear in a few minutes.' -SmtpServer mail.scotcourts.local
            $AddMBSendBehalfOfForm.Close()
            $AddMBSendBehalfOfForm.Dispose()
            Return SharedMailboxManagementForm
        }
    }
}
#########################################################################
###    Completed - Create SubForm Add Mailbox Send On Behalf Of       ###
#########################################################################
#endregion SendOB
#
#########################################################################
###########      Create SubForm Check Full Access             ###########
#########################################################################
#region CheckFull
Function CheckMBFullAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $CheckMBFullForm = New-Object System.Windows.Forms.Form
    $CheckMBFullForm.width = 745
    $CheckMBFullForm.height = 475
    $CheckMBFullForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $CheckMBFullForm.Controlbox = $false
    $CheckMBFullForm.Icon = $Icon
    $CheckMBFullForm.FormBorderStyle = 'Fixed3D'
    $CheckMBFullForm.Text = 'Mailbox - Check Full Access Permissions on Mailbox.'
    $CheckMBFullForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $CheckMBFullBox1 = New-Object System.Windows.Forms.GroupBox
    $CheckMBFullBox1.Location = '40,20'
    $CheckMBFullBox1.size = '650,125'
    $CheckMBFullBox1.text = '1. Select a MailBoxName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $CheckMBFulltextLabel2 = New-Object System.Windows.Forms.Label
    $CheckMBFulltextLabel2.Location = '20,50'
    $CheckMBFulltextLabel2.size = '200,40'
    $CheckMBFulltextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $CheckMBFullMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $CheckMBFullMBNameComboBox2.Location = '275,45'
    $CheckMBFullMBNameComboBox2.Size = '350, 350'
    $CheckMBFullMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $CheckMBFullMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $CheckMBFullMBNameComboBox2.Sorted = $true;
    $CheckMBFullMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $CheckMBFullMBNameComboBox2.DataSource = $Sharedmailboxes
    $CheckMBFullMBNameComboBox2.Add_SelectedIndexChanged( { $CheckMBFullSelectedMailBoxNametextLabel6.Text = "$($CheckMBFullMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $CheckMBFullBox2 = New-Object System.Windows.Forms.GroupBox
    $CheckMBFullBox2.Location = '40,170'
    $CheckMBFullBox2.size = '650,125'
    $CheckMBFullBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $CheckMBFulltextLabel3 = New-Object System.Windows.Forms.Label
    $CheckMBFulltextLabel3.Location = '40,40'
    $CheckMBFulltextLabel3.size = '400,40'
    $CheckMBFulltextLabel3.Text = 'Check Full Access permissions on Mailbox:' 
    $CheckMBFullSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $CheckMBFullSelectedMailBoxNametextLabel6.Location = '100,80'
    $CheckMBFullSelectedMailBoxNametextLabel6.Size = '400,40'
    $CheckMBFullSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $CheckMBFullBox3 = New-Object System.Windows.Forms.GroupBox
    $CheckMBFullBox3.Location = '40,320'
    $CheckMBFullBox3.size = '650,30'
    $CheckMBFullBox3.text = '3. Click Ok to Check Mailbox Full Access permissions or Cancel:'
    $CheckMBFullBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $CheckMBFullForm.Close()
            $CheckMBFullForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $CheckMBFullForm.Controls.AddRange(@($CheckMBFullBox1, $CheckMBFullBox2, $CheckMBFullBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $CheckMBFullBox1.Controls.AddRange(@($CheckMBFulltextLabel2, $CheckMBFullMBNameComboBox2))
    $CheckMBFullBox2.Controls.AddRange(@($CheckMBFulltextLabel3, $CheckMBFulltextLabel5, $CheckMBFullSelectedMailBoxNametextLabel6))
    #### Activate the form ###
    $CheckMBFullForm.Add_Shown( { $CheckMBFullForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckMBFullForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($CheckMBFullSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox -Check Full Access Permissions on Mailbox.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckMBFullForm.Close()
            $CheckMBFullForm.Dispose()
            break
        }
        #####################################################################
        ############           Check Full Access              ###############
        #####################################################################
        $MailboxName = Get-Mailbox -identity $CheckMBFullSelectedMailBoxNametextLabel6.Text
        $status = Get-MailboxPermission $MailboxName.Name | Where-Object { $_.AccessRights -eq 'FullAccess' } | Where-Object { $_.user -notlike "s-1*" -and @("scs\domain admins", "scs\enterprise admins", "nt authority\system", "scs\organization management") -notcontains $_.User } |
        Select-Object User, AccessRights
        $status | Select-Object @{label = 'User'; expression = { $_.User -replace '^SCS\\' } } | Sort-Object user | Out-GridView -Title "List of Users with Full Access permissions on $MailboxName mailbox" -Wait 
        $CheckMBFullForm.Close()
        $CheckMBFullForm.Dispose()
        Return SharedMailboxManagementForm
    }
}
#########################################################################
########   Completed -   Create SubForm Check Full Access     ###########
#########################################################################
#endregion CheckFull
#
#########################################################################
########      Create SubForm Check Send On Behalf Of           ##########
#########################################################################
#region CheckSendOB
Function CheckMBSendOnBehalfPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $CheckMBSendOnBehalfForm = New-Object System.Windows.Forms.Form
    $CheckMBSendOnBehalfForm.width = 745
    $CheckMBSendOnBehalfForm.height = 475
    $CheckMBSendOnBehalfForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $CheckMBSendOnBehalfForm.Controlbox = $false
    $CheckMBSendOnBehalfForm.Icon = $Icon
    $CheckMBSendOnBehalfForm.FormBorderStyle = 'Fixed3D'
    $CheckMBSendOnBehalfForm.Text = 'Mailbox - Check Send On Behalf Of permissions on Mailbox.'
    $CheckMBSendOnBehalfForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $CheckMBSendOnBehalfBox1 = New-Object System.Windows.Forms.GroupBox
    $CheckMBSendOnBehalfBox1.Location = '40,20'
    $CheckMBSendOnBehalfBox1.size = '650,125'
    $CheckMBSendOnBehalfBox1.text = '1. Select a MailBoxName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $CheckMBSendOnBehalftextLabel2 = New-Object System.Windows.Forms.Label
    $CheckMBSendOnBehalftextLabel2.Location = '20,50'
    $CheckMBSendOnBehalftextLabel2.size = '200,40'
    $CheckMBSendOnBehalftextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $CheckMBSendOnBehalfMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $CheckMBSendOnBehalfMBNameComboBox2.Location = '275,45'
    $CheckMBSendOnBehalfMBNameComboBox2.Size = '350, 350'
    $CheckMBSendOnBehalfMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $CheckMBSendOnBehalfMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $CheckMBSendOnBehalfMBNameComboBox2.Sorted = $true;
    $CheckMBSendOnBehalfMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $CheckMBSendOnBehalfMBNameComboBox2.DataSource = $Sharedmailboxes
    $CheckMBSendOnBehalfMBNameComboBox2.Add_SelectedIndexChanged( { $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.Text = "$($CheckMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $CheckMBSendOnBehalfBox2 = New-Object System.Windows.Forms.GroupBox
    $CheckMBSendOnBehalfBox2.Location = '40,170'
    $CheckMBSendOnBehalfBox2.size = '650,125'
    $CheckMBSendOnBehalfBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $CheckMBSendOnBehalftextLabel3 = New-Object System.Windows.Forms.Label
    $CheckMBSendOnBehalftextLabel3.Location = '40,40'
    $CheckMBSendOnBehalftextLabel3.size = '400,40'
    $CheckMBSendOnBehalftextLabel3.Text = 'Check Send On Behalf Of permissions on Mailbox:' 
    $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.Location = '100,80'
    $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.Size = '400,40'
    $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $CheckMBSendOnBehalfBox3 = New-Object System.Windows.Forms.GroupBox
    $CheckMBSendOnBehalfBox3.Location = '40,320'
    $CheckMBSendOnBehalfBox3.size = '650,30'
    $CheckMBSendOnBehalfBox3.text = '3. Click Ok to Check Mailbox Send On Behalf Of permissions or Cancel:'
    $CheckMBSendOnBehalfBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $CheckMBSendOnBehalfForm.Close()
            $CheckMBSendOnBehalfForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $CheckMBSendOnBehalfForm.Controls.AddRange(@($CheckMBSendOnBehalfBox1, $CheckMBSendOnBehalfBox2, $CheckMBSendOnBehalfBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $CheckMBSendOnBehalfBox1.Controls.AddRange(@($CheckMBSendOnBehalftextLabel2, $CheckMBSendOnBehalfMBNameComboBox2))
    $CheckMBSendOnBehalfBox2.Controls.AddRange(@($CheckMBSendOnBehalftextLabel3, $CheckMBSendOnBehalftextLabel5, $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $CheckMBSendOnBehalfForm.AcceptButton = $OKButton
    $CheckMBSendOnBehalfForm.CancelButton = $CancelButton
    #### Activate the form ###
    $CheckMBSendOnBehalfForm.Add_Shown( { $CheckMBSendOnBehalfForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckMBSendOnBehalfForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Check Send On Behalf Of permissions on Mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckMBSendOnBehalfForm.Close()
            $CheckMBSendOnBehalfForm.Dispose()
            break
        }
        #####################################################################
        ########       Check Send On Behalf Of Access           #############
        #####################################################################
        $MailboxName = Get-Mailbox -identity $CheckMBSendOnBehalfSelectedMailBoxNametextLabel6.Text
        $Status = Get-Mailbox $MailboxName.Name | Select-Object -ExpandProperty GrantSendOnBehalfTo
        $status -replace "^.*/", "" -replace "}", "" | Sort-Object | Out-GridView -Title "List of Users with Send On Behalf Of permissions on $MailboxName mailbox" -Wait
        $CheckMBSendOnBehalfForm.Close()
        $CheckMBSendOnBehalfForm.Dispose()
        Return SharedMailboxManagementForm
    }
}
#########################################################################
######    Completed - Create SubForm Check Send On Behalf Of      #######
#########################################################################
#endregion CheckSendOB
#
#########################################################################
######          Create SubForm Remove Full access               #########
#########################################################################
#region RemoveFull
Function RemoveMBFullAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $RemoveMBFullForm = New-Object System.Windows.Forms.Form
    $RemoveMBFullForm.width = 745
    $RemoveMBFullForm.height = 475
    $RemoveMBFullForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RemoveMBFullForm.Controlbox = $false
    $RemoveMBFullForm.Icon = $Icon
    $RemoveMBFullForm.FormBorderStyle = 'Fixed3D'
    $RemoveMBFullForm.Text = 'Mailbox - Remove Full Access Permissions for a User.'
    $RemoveMBFullForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $RemoveMBFullBox1 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBFullBox1.Location = '40,20'
    $RemoveMBFullBox1.size = '650,125'
    $RemoveMBFullBox1.text = '1. Select a UserName and MailBoxName from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $RemoveMBFulltextLabel1 = New-Object System.Windows.Forms.Label
    $RemoveMBFulltextLabel1.Location = '20,40'
    $RemoveMBFulltextLabel1.size = '100,40'
    $RemoveMBFulltextLabel1.Text = 'UserName:' 
    $RemoveMBFulltextLabel2 = New-Object System.Windows.Forms.Label
    $RemoveMBFulltextLabel2.Location = '20,80'
    $RemoveMBFulltextLabel2.size = '100,40'
    $RemoveMBFulltextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $RemoveMBFullUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $RemoveMBFullUserNameComboBox1.Location = '275,35'
    $RemoveMBFullUserNameComboBox1.Size = '350, 310'
    $RemoveMBFullUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $RemoveMBFullUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $RemoveMBFullUserNameComboBox1.Sorted = $true;
    $RemoveMBFullUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveMBFullUserNameComboBox1.DataSource = $UsernameList
    $RemoveMBFullUserNameComboBox1.Add_SelectedIndexChanged( { $RemoveMBFullSelectedUserNametextLabel4.Text = "$($RemoveMBFullUserNameComboBox1.SelectedItem.ToString())" })
    $RemoveMBFullMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $RemoveMBFullMBNameComboBox2.Location = '275,75'
    $RemoveMBFullMBNameComboBox2.Size = '350, 350'
    $RemoveMBFullMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $RemoveMBFullMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $RemoveMBFullMBNameComboBox2.Sorted = $true;
    $RemoveMBFullMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveMBFullMBNameComboBox2.DataSource = $Sharedmailboxes 
    $RemoveMBFullMBNameComboBox2.Add_SelectedIndexChanged( { $RemoveMBFullSelectedMailBoxNametextLabel6.Text = "$($RemoveMBFullMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $RemoveMBFullBox2 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBFullBox2.Location = '40,170'
    $RemoveMBFullBox2.size = '650,125'
    $RemoveMBFullBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $RemoveMBFulltextLabel3 = New-Object System.Windows.Forms.Label
    $RemoveMBFulltextLabel3.Location = '40,40'
    $RemoveMBFulltextLabel3.size = '100,40'
    $RemoveMBFulltextLabel3.Text = 'The User:' 
    $RemoveMBFullSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $RemoveMBFullSelectedUserNametextLabel4.Location = '30,80'
    $RemoveMBFullSelectedUserNametextLabel4.Size = '200,40'
    $RemoveMBFullSelectedUserNametextLabel4.ForeColor = 'Blue'
    $RemoveMBFulltextLabel5 = New-Object System.Windows.Forms.Label
    $RemoveMBFulltextLabel5.Location = '175,40'
    $RemoveMBFulltextLabel5.size = '450,40'
    $RemoveMBFulltextLabel5.Text = 'Will have Full Access permissions removed from:'
    $RemoveMBFullSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $RemoveMBFullSelectedMailBoxNametextLabel6.Location = '350,80'
    $RemoveMBFullSelectedMailBoxNametextLabel6.Size = '200,40'
    $RemoveMBFullSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $RemoveMBFullBox3 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBFullBox3.Location = '40,320'
    $RemoveMBFullBox3.size = '650,30'
    $RemoveMBFullBox3.text = '3. Click Continue to remove Full Access permissions or Cancel:'
    $RemoveMBFullBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $RemoveMBFullForm.Close()
            $RemoveMBFullForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $RemoveMBFullForm.Controls.AddRange(@($RemoveMBFullBox1, $RemoveMBFullBox2, $RemoveMBFullBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $RemoveMBFullBox1.Controls.AddRange(@($RemoveMBFulltextLabel1, $RemoveMBFulltextLabel2, $RemoveMBFullUserNameComboBox1, $RemoveMBFullMBNameComboBox2))
    $RemoveMBFullBox2.Controls.AddRange(@($RemoveMBFulltextLabel3, $RemoveMBFullSelectedUserNametextLabel4, $RemoveMBFulltextLabel5, $RemoveMBFullSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $RemoveMBFullForm.AcceptButton = $OKButton
    $RemoveMBFullForm.CancelButton = $CancelButton
    #### Activate the form ###
    $RemoveMBFullForm.add_Shown( { $RemoveMBFullForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $RemoveMBFullForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($RemoveMBFullSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Remove Full Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBFullForm.Close()
            $RemoveMBFullForm.Dispose()
            break
        }
        Elseif ($RemoveMBFullSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Shared mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Remove Full Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBFullForm.Close()
            $RemoveMBFullForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($RemoveMBFullUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailBoxPrimarySMTPAddress = get-mailbox $($RemoveMBFullMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        #########        Check if User already has full access  #############
        #####################################################################
        $Status = Get-Mailbox $MailBoxPrimarySMTPAddress | Get-MailboxPermission -User $UserSamAccountName
        If ($Status.AccessRights -ne 'FullAccess') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ($($RemoveMBFullUserNameComboBox1.SelectedItem.ToString())) does not have Full Access Permissions to the ($($RemoveMBFullMBNameComboBox2.SelectedItem.ToString())) mailbox.", 'Mailbox - Remove Full Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBFullForm.Close()
            $RemoveMBFullForm.Dispose()
            Return SharedMailboxManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######            Remove Full access for user        ################
            #####################################################################
            Remove-MailboxPermission –Identity $MailBoxPrimarySMTPAddress -User $UserSamAccountName -AccessRights FullAccess -confirm:$false
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ($($RemoveMBFullUserNameComboBox1.SelectedItem.ToString())) has had Full Access Permissions removed on the ($($RemoveMBFullMBNameComboBox2.SelectedItem.ToString())) mailbox.", "Mailbox - Remove Full Access Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($RemoveMBFullUserNameComboBox1.SelectedItem.ToString()) has had Full Access permissions removed on the $($RemoveMBFullMBNameComboBox2.SelectedItem.ToString()) mailbox" -Body "The user needs to close and re-open Outlook." -SmtpServer mail.scotcourts.local
            $RemoveMBFullForm.Close()
            $RemoveMBFullForm.Dispose()
            SharedMailboxManagementForm
        }
    }
}
#########################################################################
######     Completed - Create SubForm Remove Full access        #########
#########################################################################
#endregion RemoveFull
#
#########################################################################
####          Create SubForm Remove Send On Behalf Of          ##########
#########################################################################
#region RemoveSendOB
Function RemoveMBSendOnBehalfPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $RemoveMBSendOnBehalfForm = New-Object System.Windows.Forms.Form
    $RemoveMBSendOnBehalfForm.width = 745
    $RemoveMBSendOnBehalfForm.height = 475
    $RemoveMBSendOnBehalfForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RemoveMBSendOnBehalfForm.Controlbox = $false
    $RemoveMBSendOnBehalfForm.Icon = $Icon
    $RemoveMBSendOnBehalfForm.FormBorderStyle = 'Fixed3D'
    $RemoveMBSendOnBehalfForm.Text = 'Mailbox - Remove Send On Behalf Of permissions for a User.'
    $RemoveMBSendOnBehalfForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $RemoveMBSendOnBehalfBox1 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBSendOnBehalfBox1.Location = '40,20'
    $RemoveMBSendOnBehalfBox1.size = '650,125'
    $RemoveMBSendOnBehalfBox1.text = '1. Select a UserName and MailBoxName from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $RemoveMBSendOnBehalftextLabel1 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalftextLabel1.Location = '20,40'
    $RemoveMBSendOnBehalftextLabel1.size = '100,40'
    $RemoveMBSendOnBehalftextLabel1.Text = 'UserName:' 
    $RemoveMBSendOnBehalftextLabel2 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalftextLabel2.Location = '20,80'
    $RemoveMBSendOnBehalftextLabel2.size = '100,40'
    $RemoveMBSendOnBehalftextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $RemoveMBSendOnBehalfUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $RemoveMBSendOnBehalfUserNameComboBox1.Location = '275,35'
    $RemoveMBSendOnBehalfUserNameComboBox1.Size = '350, 310'
    $RemoveMBSendOnBehalfUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $RemoveMBSendOnBehalfUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $RemoveMBSendOnBehalfUserNameComboBox1.Sorted = $true;
    $RemoveMBSendOnBehalfUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveMBSendOnBehalfUserNameComboBox1.DataSource = $UsernameList
    $RemoveMBSendOnBehalfUserNameComboBox1.Add_SelectedIndexChanged( { $RemoveMBSendOnBehalfSelectedUserNametextLabel4.Text = "$($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString())" })
    $RemoveMBSendOnBehalfMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $RemoveMBSendOnBehalfMBNameComboBox2.Location = '275,75'
    $RemoveMBSendOnBehalfMBNameComboBox2.Size = '350, 350'
    $RemoveMBSendOnBehalfMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $RemoveMBSendOnBehalfMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $RemoveMBSendOnBehalfMBNameComboBox2.Sorted = $true;
    $RemoveMBSendOnBehalfMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveMBSendOnBehalfMBNameComboBox2.DataSource = $Sharedmailboxes   
    $RemoveMBSendOnBehalfMBNameComboBox2.Add_SelectedIndexChanged( { $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6.Text = "$($RemoveMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $RemoveMBSendOnBehalfBox2 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBSendOnBehalfBox2.Location = '40,170'
    $RemoveMBSendOnBehalfBox2.size = '650,125'
    $RemoveMBSendOnBehalfBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $RemoveMBSendOnBehalftextLabel3 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalftextLabel3.Location = '40,40'
    $RemoveMBSendOnBehalftextLabel3.size = '100,40'
    $RemoveMBSendOnBehalftextLabel3.Text = 'The User:' 
    $RemoveMBSendOnBehalfSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalfSelectedUserNametextLabel4.Location = '30,80'
    $RemoveMBSendOnBehalfSelectedUserNametextLabel4.Size = '200,40'
    $RemoveMBSendOnBehalfSelectedUserNametextLabel4.ForeColor = 'Blue'
    $RemoveMBSendOnBehalftextLabel5 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalftextLabel5.Location = '175,40'
    $RemoveMBSendOnBehalftextLabel5.size = '450,40'
    $RemoveMBSendOnBehalftextLabel5.Text = 'Will have Send On Behalf Of permissions removed from:'
    $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6.Location = '350,80'
    $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6.Size = '200,40'
    $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $RemoveMBSendOnBehalfBox3 = New-Object System.Windows.Forms.GroupBox
    $RemoveMBSendOnBehalfBox3.Location = '40,320'
    $RemoveMBSendOnBehalfBox3.size = '650,30'
    $RemoveMBSendOnBehalfBox3.text = '3. Click Continue to remove Send On Behalf Of Access permissions or Cancel:'
    $RemoveMBSendOnBehalfBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $RemoveMBSendOnBehalfForm.Close()
            $RemoveMBSendOnBehalfForm.Dispose()
            Return SharedMailboxManagementForm })
    ### Add all the Form controls ### 
    $RemoveMBSendOnBehalfForm.Controls.AddRange(@($RemoveMBSendOnBehalfBox1, $RemoveMBSendOnBehalfBox2, $RemoveMBSendOnBehalfBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $RemoveMBSendOnBehalfBox1.Controls.AddRange(@($RemoveMBSendOnBehalftextLabel1, $RemoveMBSendOnBehalftextLabel2, $RemoveMBSendOnBehalfUserNameComboBox1, $RemoveMBSendOnBehalfMBNameComboBox2))
    $RemoveMBSendOnBehalfBox2.Controls.AddRange(@($RemoveMBSendOnBehalftextLabel3, $RemoveMBSendOnBehalfSelectedUserNametextLabel4, $RemoveMBSendOnBehalftextLabel5, $RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $RemoveMBSendOnBehalfForm.AcceptButton = $OKButton
    $RemoveMBSendOnBehalfForm.CancelButton = $CancelButton
    #### Activate the form ###
    $RemoveMBSendOnBehalfForm.add_Shown( { $RemoveMBSendOnBehalfForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $RemoveMBSendOnBehalfForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($RemoveMBSendOnBehalfSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Remove Send On Behalf Of permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBSendOnBehalfForm.Close()
            $RemoveMBSendOnBehalfForm.Dispose()
            break
        }
        Elseif ($RemoveMBSendOnBehalfSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Shared mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Remove Send On Behalf Of permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBSendOnBehalfForm.Close()
            $RemoveMBSendOnBehalfForm.Dispose()
            break
        }
        #####################################################################
        ########  get user sam & display name from user name:  ############## 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $UserDisplayName = get-mailbox $($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString()) | Select-Object Name | Select-Object -ExpandProperty Name
        $MailBoxPrimarySMTPAddress = get-mailbox $($RemoveMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        #########   Check if User already has SendOnBehalf access  ##########
        #####################################################################
        $Status = Get-Mailbox $MailBoxPrimarySMTPAddress | Select-Object -ExpandProperty GrantSendOnBehalfTo
        $status1 = ($Status -match "$UserDisplayName").Count
        If ($Status1 -eq "") {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString()) ) doesnt have SendOnBehalfOf permissions to the ( $($RemoveMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString()) ) mailbox.", "Mailbox - Add SendOnbehalfOf Permission for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveMBSendOnBehalfForm.Close()
            $RemoveMBSendOnBehalfForm.Dispose()
            Return SharedMailboxManagementForm
        }    
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######       Remove Send on Behalf of for user            ###########
            #####################################################################
            Set-Mailbox $MailBoxPrimarySMTPAddress -GrantSendOnBehalfTo @{remove = "$UserSamAccountName" }
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ($($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString())) has had Send On Behalf Of Permissions removed on the ($($RemoveMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString())) mailbox.", "Mailbox - Remove Full Access Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($RemoveMBSendOnBehalfUserNameComboBox1.SelectedItem.ToString()) has had Send On Behalf Of permissions removed on the $($RemoveMBSendOnBehalfMBNameComboBox2.SelectedItem.ToString()) mailbox" -Body "The user needs to close and re-open Outlook." -SmtpServer mail.scotcourts.local
            $RemoveMBSendOnBehalfForm.Close()
            $RemoveMBSendOnBehalfForm.Dispose()
            Return SharedMailboxManagementForm
        }
    }
}
#########################################################################
####    Completed -  Create SubForm Remove Send On Behalf Of   ##########
#########################################################################
#endregion RemoveSendOB
#
#########################################################################
####          Create SubForm  Shared Log On Sub Form              #######
#########################################################################
#region ShareLogon
Function SharedLogOnForm {
    ### Show Start Message  ###
    Add-Type -AssemblyName System.Windows.Forms 
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
    $StartMessage = [System.Windows.Forms.MessageBox]::Show("This script manages Shared Mailboxes.`n`nNote 1: When prompted for a username & password enter:  scs\YourLogOnName & password`n`nNote 2:When prompted for an Outlook profile add a new profile named:  SharedMailbox `n`nNote 3:When adding new profile the email address needs changed to Shared Mailbox Name  Please click OK to continue or Cancel to exit", 'Shared Mailbox - Manage Shared Mailboxes.', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($StartMessage -eq 'Cancel') { exit } 
    ### Check if Outlook running  ###
    if ($Null -ne (Get-Process "outlook" -ea SilentlyContinue)) {
        Add-Type -AssemblyName System.Windows.Forms 
        [System.Windows.Forms.MessageBox]::Show("Outlook is currently running!!!!.`n`nYou need to close Outlook before running this script.", 'Shared Mailbox - Manage Shared Mailboxes', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Return SharedMailboxManagementForm
    }
    else {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $SharedLogOnForm = New-Object System.Windows.Forms.Form
        $SharedLogOnForm.width = 745
        $SharedLogOnForm.height = 475
        $SharedLogOnForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $SharedLogOnForm.Controlbox = $false
        $SharedLogOnForm.Icon = $Icon
        $SharedLogOnForm.FormBorderStyle = 'Fixed3D'
        $SharedLogOnForm.Text = 'Shared Mailbox - Manage Shared Mailboxes v1.0'
        $SharedLogOnForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $SharedLogOnBox1 = New-Object System.Windows.Forms.GroupBox
        $SharedLogOnBox1.Location = '40,20'
        $SharedLogOnBox1.size = '650,125'
        $SharedLogOnBox1.text = '1. Select a MailBoxName from the dropdown list:'
        ### Create group 1 box text labels. ###
        $SharedLogOntextLabel2 = New-Object System.Windows.Forms.Label;
        $SharedLogOntextLabel2.Location = '20,50'
        $SharedLogOntextLabel2.size = '200,40'
        $SharedLogOntextLabel2.Text = 'MailBoxName:' 
        ### Create group 1 box combo boxes. ###
        $SharedLogOnMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
        $SharedLogOnMBNameComboBox2.Location = '275,45'
        $SharedLogOnMBNameComboBox2.Size = '350, 350'
        $SharedLogOnMBNameComboBox2.AutoCompleteMode = 'Suggest'
        $SharedLogOnMBNameComboBox2.AutoCompleteSource = 'ListItems'
        $SharedLogOnMBNameComboBox2.Sorted = $true;
        $SharedLogOnMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SharedLogOnMBNameComboBox2.DataSource = $Sharedmailboxes
        $SharedLogOnMBNameComboBox2.Add_SelectedIndexChanged( { $SharedLogOnSelectedMailBoxNametextLabel6.Text = "$($SharedLogOnMBNameComboBox2.SelectedItem.ToString())" })
        ### Create group 2 box in form. ###
        $SharedLogOnBox2 = New-Object System.Windows.Forms.GroupBox
        $SharedLogOnBox2.Location = '40,170'
        $SharedLogOnBox2.size = '650,125'
        $SharedLogOnBox2.text = '2. Check the details below are correct before proceeding:'
        ### Create group 2 box text labels. ###
        $SharedLogOntextLabel3 = New-Object System.Windows.Forms.Label;
        $SharedLogOntextLabel3.Location = '40,40'
        $SharedLogOntextLabel3.size = '400,40'
        $SharedLogOntextLabel3.Text = 'Manage Shared Mailbox:' 
        $SharedLogOnSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
        $SharedLogOnSelectedMailBoxNametextLabel6.Location = '100,80'
        $SharedLogOnSelectedMailBoxNametextLabel6.Size = '400,40'
        $SharedLogOnSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
        ### Create group 3 box in form. ###
        $SharedLogOnBox3 = New-Object System.Windows.Forms.GroupBox
        $SharedLogOnBox3.Location = '40,320'
        $SharedLogOnBox3.size = '650,30'
        $SharedLogOnBox3.text = '3. Click Ok to Manage Shared Mailbox or Cancel:'
        $SharedLogOnBox3.button
        ### Add an OK button ###
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '590,370'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '470,370'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to Form'
        $CancelButton.Add_Click( {
                $SharedLogOnForm.Close()
                $SharedLogOnForm.Dispose()
                Return SharedMailboxManagementForm })
        ### Add all the Form controls ### 
        $SharedLogOnForm.Controls.AddRange(@($SharedLogOnBox1, $SharedLogOnBox2, $SharedLogOnBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $SharedLogOnBox1.Controls.AddRange(@($SharedLogOntextLabel2, $SharedLogOnMBNameComboBox2))
        $SharedLogOnBox2.Controls.AddRange(@($SharedLogOntextLabel3, $SharedLogOntextLabel5, $SharedLogOnSelectedMailBoxNametextLabel6))
        #### Activate the form ###
        $SharedLogOnForm.Add_Shown( { $SharedLogOnForm.Activate() })    
        #### Get the results from the button click ###
        $dialogResult = $SharedLogOnForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            ### Don't accept null mailbox ### 
            if ($SharedLogOnSelectedMailBoxNametextLabel6.Text -eq '') {
                [System.Windows.Forms.MessageBox]::Show("You need to select a Mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Shared Mailbox - Manage Shared Mailboxes.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $SharedLogOnForm.Close()
                $SharedLogOnForm.Dispose()
                Return SharedMailboxManagementForm
            }
            Else {
                ###  Get mailbox primary smtpaddress from mailbox display name: ###
                $MailBoxName = get-mailbox $SharedLogOnSelectedMailBoxNametextLabel6.Text | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
                ### Get user log on name & SID  ###
                $UserName = $env:username
                $UserSamName = $UserName -replace "admin", ""
                ## removal of admin from end of user name, changed to this method for win10 as $UserName = Get-WMIObject -class Win32_ComputerSystem | Select-Object -Expandproperty UserName method blocked by GPO.
                $SID = (Get-ADUser -Identity $UserSamName).SID.Value
                ### CHECK - continue if only 1 mailbox & sid  ###
                if (($UserSamName | Measure-Object).count -ne 1) { SharedLogOn }
                if (($MailBoxName | Measure-Object).count -ne 1) { SharedLogOn }
                ### Add mailbox permission  ###
                Add-MailboxPermission $MailBoxName -User $UserSamName -AccessRights FullAccess -confirm:$false -InheritanceType All -Automapping $false
                ### if sharedmailbox profile exists remove  ###
                If ((Test-Path -Path "Registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\SharedMailbox") -eq '$true') {
                    Remove-Item -Path "Registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\SharedMailbox" -Recurse
                }
                $arguments = '-profiles'
                Start-Process -FilePath 'C:\program files (x86)\microsoft office\office16\OUTLOOK.EXE' -ArgumentList $arguments -Credential (Get-Credential) -Wait
                Remove-Variable arguments
                ### If outlook closed without creating profile exit  ###
                If ((Test-Path -Path "Registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\SharedMailbox") -ne '$true') {
                    #[System.Windows.Forms.MessageBox]::Show("Outlook has been closed without adding a SharedMailbox profile  You need to add an outlook profile to manage a shared mailbox.", "Shared Mailbox - Manage Shared Mailboxes.", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    [System.Windows.Forms.MessageBox]::Show("Outlook has been closed. the SharedMailbox profile Has NOT been deleted, Please remove this.", "Shared Mailbox - Manage Shared Mailboxes.", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    ### Remove mailbox permission  ###
                    if (($UserSamName | Measure-Object).count -ne 1) { SharedLogOn }
                    if (($MailBoxName | Measure-Object).count -ne 1) { SharedLogOn }
                    Remove-MailboxPermission $MailBoxName -User $UserSamName -AccessRights FullAccess -confirm:$false
                    $SharedLogOnForm.Close()
                    $SharedLogOnForm.Dispose()
                    Return SharedMailboxManagementForm
                }
                Else {
                    ### Remove mailbox permission  ###
                    if (($UserSamName | Measure-Object).count -ne 1) { SharedLogOn }
                    if (($MailBoxName | Measure-Object).count -ne 1) { SharedLogOn }
                    Remove-MailboxPermission $MailBoxName -User $UserSamName -AccessRights FullAccess -confirm:$false
                    ### Remove outlook profile  ###
                    If ((Test-Path -Path "Registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\SharedMailbox") -eq '$true') {
                        Remove-Item -Path "Registry::\HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Profiles\SharedMailbox" -Recurse
                    }
                    $SharedLogOnForm.Close()
                    $SharedLogOnForm.Dispose()
                    Return SharedMailboxManagementForm
                }
            }
        }
    }
}  
#########################################################################
####  Completed - Create SubForm  Shared Log On Sub Form Sub Form  ######
#########################################################################
#endregion SharedLogon
#
#########################################################################
####     Create SubForm  Add New Shared mailbox Sub Form          #######
#########################################################################
#region NewShared
Function AddNewSharedMBForm {
    #########################################################################
    #######              Show Start Message:                  ###############
    #########################################################################
    Add-Type -AssemblyName System.Windows.Forms 
    $StartMessage = [System.Windows.Forms.MessageBox]::Show("This script creates a New Shared Mailbox.  The new mailbox will be created in the Courts\Users\Shared Mailboxes OU in AD.  Please Note 1: There are 4 fields & all the fields are mandatory.  Please note 2: The Display Name will appear in the global address list and should start with a capital letter.  Please Note 3: The email address entered doesnt need @scotcourts.gov.uk added and should have no spaces or non-alpha characters.  Please click OK to continue or Cancel to exit", 'Add New Shared Mailbox.', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($StartMessage -eq 'Cancel') { exit } 
    else {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $SharedMBNewForm = New-Object System.Windows.Forms.Form
        $SharedMBNewForm.width = 780
        $SharedMBNewForm.height = 550
        $SharedMBNewForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $SharedMBNewForm.Controlbox = $false
        $SharedMBNewForm.Icon = $Icon
        $SharedMBNewForm.FormBorderStyle = 'Fixed3D'
        $SharedMBNewForm.Text = 'Add New Shared Mailbox'
        $SharedMBNewForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $SharedMBNewBox1 = New-Object System.Windows.Forms.GroupBox
        $SharedMBNewBox1.Location = '40,20'
        $SharedMBNewBox1.size = '700,200'
        $SharedMBNewBox1.text = '1. Enter the Shared Mailbox details:'
        ### Create group 1 box text labels. ###
        $SharedMBNewtextLabel1 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel1.Location = '20,35'
        $SharedMBNewtextLabel1.size = '350,40'
        $SharedMBNewtextLabel1.Text = 'Display Name:          (e.g. Haddington Fines Enquiries)' 
        $SharedMBNewtextLabel2 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel2.Location = '20,75'
        $SharedMBNewtextLabel2.size = '350,40'
        $SharedMBNewtextLabel2.Text = 'Email address:         (e.g. haddingtonfinesenquiries)' 
        $SharedMBNewtextLabel3 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel3.Location = '20,112'
        $SharedMBNewtextLabel3.size = '350,40'
        $SharedMBNewtextLabel3.Text = 'Owners name:          (e.g.Joe Bloggs)' 
        $SharedMBNewtextLabel4 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel4.Location = '20,150'
        $SharedMBNewtextLabel4.size = '370,40'
        $SharedMBNewtextLabel4.Text = 'Mailbox description: (e.g. Haddington Fines shared mailbox)' 
        ### Create group 1 box text boxes. ###
        $SharedMBNewtextBox1 = New-Object System.Windows.Forms.TextBox
        $SharedMBNewtextBox1.Location = '425,30'
        $SharedMBNewtextBox1.Size = '250,40'
        $SharedMBNewtextBox1.add_TextChanged( { $SharedMBNewtextLabel6.Text = "$($SharedMBNewtextBox1.text)" })
        $SharedMBNewtextBox1.Add_TextChanged( { If ($This.Text -and $SharedMBNewtextBox2.Text -and $SharedMBNewtextBox3.Text -and $SharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } })  
        $SharedMBNewtextBox2 = New-Object System.Windows.Forms.TextBox
        $SharedMBNewtextBox2.Location = '425,70'
        $SharedMBNewtextBox2.Size = '250,40'
        $SharedMBNewtextBox2.add_textChanged( { $SharedMBNewtextLabel8.Text = "$($SharedMBNewtextBox2.text)@scotcourts.gov.uk" })
        $SharedMBNewtextBox2.Add_TextChanged( { If ($This.Text -and $SharedMBNewtextBox1.Text -and $SharedMBNewtextBox3.Text -and $SharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } }) 
        $SharedMBNewtextBox3 = New-Object System.Windows.Forms.TextBox
        $SharedMBNewtextBox3.Location = '425,105'
        $SharedMBNewtextBox3.Size = '250,40'
        $SharedMBNewtextBox3.add_TextChanged( { $SharedMBNewtextLabel10.Text = "$($SharedMBNewtextBox3.text)" })
        $SharedMBNewtextBox3.Add_TextChanged( { If ($This.Text -and $SharedMBNewtextBox1.Text -and $SharedMBNewtextBox2.Text -and $SharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } })  
        $SharedMBNewtextBox4 = New-Object System.Windows.Forms.TextBox
        $SharedMBNewtextBox4.Location = '425,145'
        $SharedMBNewtextBox4.Size = '250,40'
        $SharedMBNewtextBox4.add_textChanged( { $SharedMBNewtextLabel12.Text = "$($SharedMBNewtextBox4.text)" })
        $SharedMBNewtextBox4.Add_TextChanged( { If ($This.Text -and $SharedMBNewtextBox1.Text -and $SharedMBNewtextBox2.Text -and $SharedMBNewtextBox3.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } }) 
        ### Create group 2 box in form. ###
        $SharedMBNewBox2 = New-Object System.Windows.Forms.GroupBox
        $SharedMBNewBox2.Location = '40,225'
        $SharedMBNewBox2.size = '700,175'
        $SharedMBNewBox2.text = '2. Check the details below are correct before proceeding:'
        ### Create group 2 box text labels.
        $SharedMBNewtextLabel5 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel5.Location = '20,30'
        $SharedMBNewtextLabel5.size = '350,30'
        $SharedMBNewtextLabel5.Text = 'Shared Mailbox will appear in Global Adress List as:' 
        $SharedMBNewtextLabel6 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel6.Location = '40,65'
        $SharedMBNewtextLabel6.Size = '250,30'
        $SharedMBNewtextLabel6.ForeColor = 'Blue'
        $SharedMBNewtextLabel7 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel7.Location = '430,30'
        $SharedMBNewtextLabel7.size = '200,30'
        $SharedMBNewtextLabel7.Text = 'With the email address:'
        $SharedMBNewtextLabel8 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel8.Location = '380,65'
        $SharedMBNewtextLabel8.Size = '400,30'
        $SharedMBNewtextLabel8.ForeColor = 'Blue'
        $SharedMBNewtextLabel9 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel9.Location = '20,95'
        $SharedMBNewtextLabel9.size = '90,30'
        $SharedMBNewtextLabel9.Text = 'The owner is:' 
        $SharedMBNewtextLabel10 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel10.Location = '40,125'
        $SharedMBNewtextLabel10.Size = '220,30'
        $SharedMBNewtextLabel10.ForeColor = 'Blue'
        $SharedMBNewtextLabel11 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel11.Location = '430,95'
        $SharedMBNewtextLabel11.size = '200,30'
        $SharedMBNewtextLabel11.Text = 'The description in AD is:'
        $SharedMBNewtextLabel12 = New-Object System.Windows.Forms.Label
        $SharedMBNewtextLabel12.Location = '380,125'
        $SharedMBNewtextLabel12.Size = '400,30'
        $SharedMBNewtextLabel12.ForeColor = 'Blue'
        ### Create group 3 box in form. ###
        $SharedMBNewBox3 = New-Object System.Windows.Forms.GroupBox
        $SharedMBNewBox3.Location = '40,410'
        $SharedMBNewBox3.size = '700,30'
        $SharedMBNewBox3.text = '3. Click Ok to add New Shared Mailbox or Cancel:'
        $SharedMBNewBox3.button
        ### Add an OK button ###
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '640,460'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '525,460'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to Form'
        $CancelButton.add_Click( {
                $SharedMBNewForm.Close()
                $SharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm })
        ### Add all the Form controls ### 
        $SharedMBNewForm.Controls.AddRange(@($SharedMBNewBox1, $SharedMBNewBox2, $SharedMBNewBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $SharedMBNewBox1.Controls.AddRange(@($SharedMBNewtextLabel1, $SharedMBNewtextLabel2, $SharedMBNewtextLabel3, $SharedMBNewtextLabel4, $SharedMBNewtextBox1, $SharedMBNewtextBox2, $SharedMBNewtextBox3, $SharedMBNewtextBox4))
        $SharedMBNewBox2.Controls.AddRange(@($SharedMBNewtextLabel5, $SharedMBNewtextLabel6, $SharedMBNewtextLabel7, $SharedMBNewtextLabel8, $SharedMBNewtextLabel9, $SharedMBNewtextLabel10, $SharedMBNewtextLabel11, $SharedMBNewtextLabel12))
        #### Assign the Accept and Cancel options in the form ### 
        $SharedMBNewForm.AcceptButton = $OKButton
        $SharedMBNewForm.CancelButton = $CancelButton
        #### Activate the form ###
        $SharedMBNewForm.Add_Shown( { $SharedMBNewForm.Activate() })    
        #### Get the results from the button click ###
        $dialogResult = $SharedMBNewForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            #####################################################################
            ########   Don't accept null username or mailbox     ################ 
            #####################################################################
            if ($SharedMBNewtextBox1.text -eq '') {
                [System.Windows.Forms.MessageBox]::Show("You need to type in details !!!!!  Trying to enter blank fields is never a good idea.", 'Add New Shared Mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $SharedMBNewForm.Close()
                $SharedMBNewForm.Dispose()
                break
            }
            #####################################################################
            #########   Check if email address is already in use    #############
            #####################################################################
            $DisplayName = $SharedMBNewtextBox1.Text
            $EmailName = $SharedMBNewtextBox2.Text   
            $ListEmailAddress = Get-ADObject -Filter "mail -eq '$EmailName@scotcourts.gov.uk'" | Measure-Object count 
            If ($Null -ne $ListEmailAddress) {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The Shared Mailbox - $DisplayName - can not be added because the email address $EmailName@scotcourts.gov.uk is currently in use on another shared mailbox  Please use a name/email address that's not currently in use.", "ERROR - CAN NOT ADD NEW SHARED MAILBOX", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $SharedMBNewForm.Close()
                $SharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm
            }
            Else {
                #####################################################################
                #    CHECK - continue if only 1 EmailName in pipe if not exit       #
                #####################################################################
                if (($EmailName | Measure-Object).count -ne 1) { AddNewSharedMBForm }
                #####################################################################
                # 
                #####################################################################
                ######           Add New Shared mailbox               ###############
                #####################################################################
                New-Mailbox -shared  -Name "$DisplayName" -Alias $EmailName -UserprincipalName "$EmailName@scotcourts.local" -OrganizationalUnit 'ou=shared mailboxes,ou=resource accounts,ou=useraccounts,ou=courts,dc=scotcourts,dc=local'
                # 
                ##########################################################################
                #######   Disable Pop, OWA, Imap & ActiveSync for user ###################
                ##########################################################################
                Set-CASMailbox -Identity $EmailName -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False 
                #  
                #####################################################################
                ##################  Message complete message  #######################
                #####################################################################
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("New Shared Mailbox - $DisplayName - has been added.  The users who need:  Full Access and  Send on Behalf Of permissions  now need to be added to the new mailbox.", 'Add New Shared Mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                #####################################################################
                #############   Reload Shared Mailbox lists       ###################
                #####################################################################
                ###  Create form to pause for 7 sec  ###
                ########################################
                Add-Type -AssemblyName System.Windows.Forms
                ### Build Form ###
                $objForm = New-Object System.Windows.Forms.Form
                $objForm.Text = 'Add New Shared Mailbox'
                $objForm.Size = New-Object System.Drawing.Size(450, 270)
                $objForm.StartPosition = 'CenterScreen'
                $objForm.Controlbox = $false
                #### Add Label ###
                $objLabel = New-Object System.Windows.Forms.Label
                $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
                $objLabel.Size = New-Object System.Drawing.Size(300, 120)
                $objLabel.Text = 'The Shared Mailbox List is being reloaded from AD   with the new mailbox you have just added.     Please Wait. ..............'
                $objForm.Controls.Add($objLabel)
                ### Show the form ###
                $objForm.Show() | Out-Null
                ### wait 7 seconds ###
                Start-Sleep -Seconds 7
                ### destroy form ###
                $objForm.Close() | Out-Null
                ### update shared mailbox lists with new added mailbox ###
                function changeit {  
                    $script:Sharedmailboxes = Get-Mailbox -OrganizationalUnit 'ou=shared mailboxes,ou=resource accounts,ou=useraccounts,ou=courts,dc=scotcourts,dc=local' | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName 
                }  
                changeit 
                ##########################################################################
                #######            Add description & mailbox  owner    ###################
                ##########################################################################
                Set-ADUser –PassThru -Identity $EmailName -Office "Owner: $($SharedMBNewtextLabel10.text)"
                Set-ADUser -Identity $EmailName -Description $SharedMBNewtextLabel12.Text
                #####################################################################
                #############   Send email to helpdesk    ###########################
                #####################################################################
                Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New Shared Mailbox - $DisplayName - has been added" -Body "New Shared mailbox - $DisplayName - has been added  The users who need:  Full Access and  Send on Behalf Of permissions  now need to be added to the new mailbox." -SmtpServer mail.scotcourts.local
                $SharedMBNewForm.Close()
                $SharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm
            }
        }
    }
}
#########################################################################
#####    Completed -  Create Add New Shared mailbox Sub Form        #####
######################################################################### 
#endregion NewShared
#
#########################################################################
###      Create SubForm  Add New Tribunal Shared mailbox Sub Form   #####
#########################################################################
#region NewTribShared
Function AddNewTribunalSharedMBForm {
    #########################################################################
    #######              Show Start Message:                  ###############
    #########################################################################
    Add-Type -AssemblyName System.Windows.Forms 
    $StartMessage = [System.Windows.Forms.MessageBox]::Show("This script creates a New Tribunal Shared mailbox.  The new mailbox will be created in the Courts\Users\Shared mailboxes OU in AD.  Please Note 1: There are 4 fields & all the fields are mandatory.  Please note 2: The Display Name will appear in the global address list and should start with a capital letter.  Please Note 3: The email address entered doesnt need @scotcourts.gov.uk added and should have no spaces or non-alpha characters.  Please click OK to continue or Cancel to exit", 'Add New Tribunal Shared mailbox.', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($StartMessage -eq 'Cancel') { exit } 
    else {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $TribSharedMBNewForm = New-Object System.Windows.Forms.Form
        $TribSharedMBNewForm.width = 780
        $TribSharedMBNewForm.height = 550
        $TribSharedMBNewForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $TribSharedMBNewForm.Controlbox = $false
        $TribSharedMBNewForm.Icon = $Icon
        $TribSharedMBNewForm.FormBorderStyle = 'Fixed3D'
        $TribSharedMBNewForm.Text = 'Add New Tribunal Shared mailbox'
        $TribSharedMBNewForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $TribSharedMBNewBox1 = New-Object System.Windows.Forms.GroupBox
        $TribSharedMBNewBox1.Location = '40,20'
        $TribSharedMBNewBox1.size = '700,200'
        $TribSharedMBNewBox1.text = '1. Enter the Tribunal Shared mailbox details:'
        ### Create group 1 box text labels. ###
        $TribSharedMBNewtextLabel1 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel1.Location = '20,35'
        $TribSharedMBNewtextLabel1.size = '350,40'
        $TribSharedMBNewtextLabel1.Text = 'Display Name:          (e.g. MHTS Hearings)' 
        $TribSharedMBNewtextLabel2 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel2.Location = '20,75'
        $TribSharedMBNewtextLabel2.size = '350,40'
        $TribSharedMBNewtextLabel2.Text = 'Email address:         (e.g. mhtshearings)' 
        $TribSharedMBNewtextLabel3 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel3.Location = '20,112'
        $TribSharedMBNewtextLabel3.size = '350,40'
        $TribSharedMBNewtextLabel3.Text = 'Owners name:          (e.g.Joe Bloggs)' 
        $TribSharedMBNewtextLabel4 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel4.Location = '20,150'
        $TribSharedMBNewtextLabel4.size = '370,40'
        $TribSharedMBNewtextLabel4.Text = 'Mailbox description: (e.g. MHTS Hearings Shared mailbox)' 
        ### Create group 1 box text boxes. ###
        $TribSharedMBNewtextBox1 = New-Object System.Windows.Forms.TextBox
        $TribSharedMBNewtextBox1.Location = '425,30'
        $TribSharedMBNewtextBox1.Size = '250,40'
        $TribSharedMBNewtextBox1.add_TextChanged( { $TribSharedMBNewtextLabel6.Text = "$($TribSharedMBNewtextBox1.text)" })
        $TribSharedMBNewtextBox1.Add_TextChanged( { If ($This.Text -and $TribSharedMBNewtextBox2.Text -and $TribSharedMBNewtextBox3.Text -and $TribSharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } })  
        $TribSharedMBNewtextBox2 = New-Object System.Windows.Forms.TextBox
        $TribSharedMBNewtextBox2.Location = '425,70'
        $TribSharedMBNewtextBox2.Size = '250,40'
        $TribSharedMBNewtextBox2.add_textChanged( { $TribSharedMBNewtextLabel8.Text = "$($TribSharedMBNewtextBox2.text)@scotcourtstribunals.gov.uk" })
        $TribSharedMBNewtextBox2.Add_TextChanged( { If ($This.Text -and $TribSharedMBNewtextBox1.Text -and $TribSharedMBNewtextBox3.Text -and $TribSharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } }) 
        $TribSharedMBNewtextBox3 = New-Object System.Windows.Forms.TextBox
        $TribSharedMBNewtextBox3.Location = '425,105'
        $TribSharedMBNewtextBox3.Size = '250,40'
        $TribSharedMBNewtextBox3.add_TextChanged( { $TribSharedMBNewtextLabel10.Text = "$($TribSharedMBNewtextBox3.text)" })
        $TribSharedMBNewtextBox3.Add_TextChanged( { If ($This.Text -and $TribSharedMBNewtextBox1.Text -and $TribSharedMBNewtextBox2.Text -and $TribSharedMBNewtextBox4.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } })  
        $TribSharedMBNewtextBox4 = New-Object System.Windows.Forms.TextBox
        $TribSharedMBNewtextBox4.Location = '425,145'
        $TribSharedMBNewtextBox4.Size = '250,40'
        $TribSharedMBNewtextBox4.add_textChanged( { $TribSharedMBNewtextLabel12.Text = "$($TribSharedMBNewtextBox4.text)" })
        $TribSharedMBNewtextBox4.Add_TextChanged( { If ($This.Text -and $TribSharedMBNewtextBox1.Text -and $TribSharedMBNewtextBox2.Text -and $TribSharedMBNewtextBox3.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } }) 
        ### Create group 2 box in form. ###
        $TribSharedMBNewBox2 = New-Object System.Windows.Forms.GroupBox
        $TribSharedMBNewBox2.Location = '40,225'
        $TribSharedMBNewBox2.size = '700,175'
        $TribSharedMBNewBox2.text = '2. Check the details below are correct before proceeding:'
        ### Create group 2 box text labels.
        $TribSharedMBNewtextLabel5 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel5.Location = '20,30'
        $TribSharedMBNewtextLabel5.size = '350,30'
        $TribSharedMBNewtextLabel5.Text = 'Tribunal Shared mailbox will appear in Global Adress List as:' 
        $TribSharedMBNewtextLabel6 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel6.Location = '40,65'
        $TribSharedMBNewtextLabel6.Size = '250,30'
        $TribSharedMBNewtextLabel6.ForeColor = 'Blue'
        $TribSharedMBNewtextLabel7 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel7.Location = '430,30'
        $TribSharedMBNewtextLabel7.size = '200,30'
        $TribSharedMBNewtextLabel7.Text = 'With the email address:'
        $TribSharedMBNewtextLabel8 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel8.Location = '380,65'
        $TribSharedMBNewtextLabel8.Size = '400,30'
        $TribSharedMBNewtextLabel8.ForeColor = 'Blue'
        $TribSharedMBNewtextLabel9 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel9.Location = '20,95'
        $TribSharedMBNewtextLabel9.size = '90,30'
        $TribSharedMBNewtextLabel9.Text = 'The owner is:' 
        $TribSharedMBNewtextLabel10 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel10.Location = '40,125'
        $TribSharedMBNewtextLabel10.Size = '220,30'
        $TribSharedMBNewtextLabel10.ForeColor = 'Blue'
        $TribSharedMBNewtextLabel11 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel11.Location = '430,95'
        $TribSharedMBNewtextLabel11.size = '200,30'
        $TribSharedMBNewtextLabel11.Text = 'The description in AD is:'
        $TribSharedMBNewtextLabel12 = New-Object System.Windows.Forms.Label
        $TribSharedMBNewtextLabel12.Location = '380,125'
        $TribSharedMBNewtextLabel12.Size = '400,30'
        $TribSharedMBNewtextLabel12.ForeColor = 'Blue'
        ### Create group 3 box in form. ###
        $TribSharedMBNewBox3 = New-Object System.Windows.Forms.GroupBox
        $TribSharedMBNewBox3.Location = '40,410'
        $TribSharedMBNewBox3.size = '700,30'
        $TribSharedMBNewBox3.text = '3. Click Ok to add New Tribunal Shared mailbox or Cancel:'
        $TribSharedMBNewBox3.button
        ### Add an OK button ###
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '640,460'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '525,460'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to Form'
        $CancelButton.add_Click( {
                $TribSharedMBNewForm.Close()
                $TribSharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm })
        ### Add all the Form controls ### 
        $TribSharedMBNewForm.Controls.AddRange(@($TribSharedMBNewBox1, $TribSharedMBNewBox2, $TribSharedMBNewBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $TribSharedMBNewBox1.Controls.AddRange(@($TribSharedMBNewtextLabel1, $TribSharedMBNewtextLabel2, $TribSharedMBNewtextLabel3, $TribSharedMBNewtextLabel4, $TribSharedMBNewtextBox1, $TribSharedMBNewtextBox2, $TribSharedMBNewtextBox3, $TribSharedMBNewtextBox4))
        $TribSharedMBNewBox2.Controls.AddRange(@($TribSharedMBNewtextLabel5, $TribSharedMBNewtextLabel6, $TribSharedMBNewtextLabel7, $TribSharedMBNewtextLabel8, $TribSharedMBNewtextLabel9, $TribSharedMBNewtextLabel10, $TribSharedMBNewtextLabel11, $TribSharedMBNewtextLabel12))
        #### Assign the Accept and Cancel options in the form ### 
        $TribSharedMBNewForm.AcceptButton = $OKButton
        $TribSharedMBNewForm.CancelButton = $CancelButton
        #### Activate the form ###
        $TribSharedMBNewForm.Add_Shown( { $TribSharedMBNewForm.Activate() })    
        #### Get the results from the button click ###
        $dialogResult = $TribSharedMBNewForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            #####################################################################
            ########   Don't accept null username or mailbox     ################ 
            #####################################################################
            if ($TribSharedMBNewtextBox1.text -eq '') {
                [System.Windows.Forms.MessageBox]::Show("You need to type in details !!!!!  Trying to enter blank fields is never a good idea.", 'Add New Tribunal Shared mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $TribSharedMBNewForm.Close()
                $TribSharedMBNewForm.Dispose()
                break
            }
            #####################################################################
            #########   Check if email address is already in use    #############
            #####################################################################
            $DisplayName = $TribSharedMBNewtextBox1.Text
            $EmailName = $TribSharedMBNewtextBox2.Text   
            $ListEmailAddress = Get-ADObject -Filter "mail -eq '$EmailName@scotcourts.gov.uk'" | Measure-Object count 
            If ($Null -ne $ListEmailAddress) {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The Tribunal Shared mailbox - $DisplayName - can not be added because the email address $EmailName@scotcourts.gov.uk is currently in use on another shared mailbox  Please use a name/email address that's not currently in use.", 'ERROR - CAN NOT ADD NEW Tribunal Shared mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $TribSharedMBNewForm.Close()
                $TribSharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm
            }
            Else {
                #####################################################################
                #    CHECK - continue if only 1 EmailName in pipe if not exit       #
                #####################################################################
                if (($EmailName | Measure-Object).count -ne 1) { AddNewSharedMBForm }
                #####################################################################
                # 
                #####################################################################
                ####       Add New Tribunal Shared mailbox              #############
                #####################################################################
                New-Mailbox -shared  -Name "$DisplayName" -Alias $EmailName -UserprincipalName "$EmailName@scotcourts.local" -OrganizationalUnit 'ou=shared mailboxes,ou=resource accounts,ou=useraccounts,ou=courts,dc=scotcourts,dc=local'
                # 
                ##########################################################################
                #######   Disable Pop, OWA, Imap & ActiveSync for user ###################
                ##########################################################################
                Set-CASMailbox -Identity $EmailName -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False 
                #
                ##########################################################################
                ####  Set Custom Attribute to apply policy to use scotcourtstribunal  ####
                ##########################################################################
                Set-Mailbox -Identity $EmailName -CustomAttribute1 'Tribunal Shared Mailbox'
                #      
                #####################################################################
                ##################  Message complete message  #######################
                #####################################################################
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("New Tribunal Shared mailbox - $DisplayName - has been added.  The users who need:  Full Access and  Send on Behalf Of permissions  now need to be added to the new mailbox.", 'Add New Tribunal Shared mailbox', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                #####################################################################
                #############   Reload Shared Mailbox lists       ###################
                #####################################################################
                ###  Create form to pause for 7 sec  ###
                ########################################
                Add-Type -AssemblyName System.Windows.Forms
                ### Build Form ###
                $objForm = New-Object System.Windows.Forms.Form
                $objForm.Text = 'Add New Tribunal Shared mailbox'
                $objForm.Size = New-Object System.Drawing.Size(450, 270)
                $objForm.StartPosition = 'CenterScreen'
                $objForm.Controlbox = $false
                #### Add Label ###
                $objLabel = New-Object System.Windows.Forms.Label
                $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
                $objLabel.Size = New-Object System.Drawing.Size(300, 120)
                $objLabel.Text = 'The Tribunal Shared mailbox List is being reloaded from AD   with the new mailbox you have just added.     Please Wait. ..............'
                $objForm.Controls.Add($objLabel)
                ### Show the form ###
                $objForm.Show() | Out-Null
                ### wait 7 seconds ###
                Start-Sleep -Seconds 7
                ### destroy form ###
                $objForm.Close() | Out-Null
                ### update Tribunal Shared mailbox lists with new added mailbox ###
                function changeit {  
                    $script:Sharedmailboxes = Get-Mailbox -OrganizationalUnit 'ou=shared mailboxes,ou=resource accounts,ou=useraccounts,ou=courts,dc=scotcourts,dc=local' | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName 
                }  
                changeit 
                ##########################################################################
                #######            Add description & mailbox  owner    ###################
                ##########################################################################
                Set-ADUser –PassThru -Identity $EmailName -Office "Owner: $($TribSharedMBNewtextLabel10.text)"
                Set-ADUser -Identity $EmailName -Description $TribSharedMBNewtextLabel12.Text
                #    
                #####################################################################
                #############   Send email to helpdesk    ###########################
                #####################################################################
                Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New Tribunal Shared mailbox - $DisplayName - has been added" -Body "New Tribunal Shared mailbox - $DisplayName - has been added  The users who need:  Full Access and  Send on Behalf Of permissions  now need to be added to the new mailbox." -SmtpServer mail.scotcourts.local
                $TribSharedMBNewForm.Close()
                $TribSharedMBNewForm.Dispose()
                Return SharedMailboxManagementForm
            }
        }
    }
}
#########################################################################
###    Completed -  Create Add New Tribunal Shared mailbox Sub Form  ####
######################################################################### 
#endregion NewTribShared
#
#########################################################################
####             Create SubForm  Copy to sent items               #######
#########################################################################
#region Copy
Function CopySharedMBForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $Copyform = New-Object System.Windows.Forms.Form
    $Copyform.width = 780
    $Copyform.height = 500
    $Copyform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Copyform.Controlbox = $false
    $Copyform.Icon = $Icon
    $Copyform.FormBorderStyle = 'Fixed3D'
    $Copyform.Text = 'Add copy of sent mail to shared mailbox sent folder form.'
    $Copyform.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $CopytBox1 = New-Object System.Windows.Forms.GroupBox
    $CopytBox1.Location = '40,40'
    $CopytBox1.size = '700,125'
    $CopytBox1.text = '1. Select a Shared Mailbox from the dropdown list:'
    ### Create group 1 box text labels. ###
    $CopyttextLabel1 = New-Object System.Windows.Forms.Label
    $CopyttextLabel1.Location = '20,40'
    $CopyttextLabel1.size = '150,40'
    $CopyttextLabel1.Text = 'Shared Mailbox:' 
    ### Create group 1 box combo boxes. ###
    $CopyMailboxComboBox1 = New-Object System.Windows.Forms.ComboBox
    $CopyMailboxComboBox1.Location = '325,35'
    $CopyMailboxComboBox1.Size = '350, 310'
    $CopyMailboxComboBox1.AutoCompleteMode = 'Suggest'
    $CopyMailboxComboBox1.AutoCompleteSource = 'ListItems'
    $CopyMailboxComboBox1.Sorted = $true;
    $CopyMailboxComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $CopyMailboxComboBox1.DataSource = $Sharedmailboxes
    $CopyMailboxComboBox1.add_SelectedIndexChanged( { $CopytSelectedMailboxtextLabel4.Text = "$($CopyMailboxComboBox1.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $CopytBox2 = New-Object System.Windows.Forms.GroupBox
    $CopytBox2.Location = '40,190'
    $CopytBox2.size = '700,125'
    $CopytBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $CopyttextLabel3 = New-Object System.Windows.Forms.Label
    $CopyttextLabel3.Location = '40,40'
    $CopyttextLabel3.size = '200,40'
    $CopyttextLabel3.Text = 'The Sent items folder of Mailbox:' 
    $CopytSelectedMailboxtextLabel4 = New-Object System.Windows.Forms.Label
    $CopytSelectedMailboxtextLabel4.Location = '250,40'
    $CopytSelectedMailboxtextLabel4.Size = '200,40'
    $CopytSelectedMailboxtextLabel4.ForeColor = 'Blue'
    $CopyttextLabel5 = New-Object System.Windows.Forms.Label
    $CopyttextLabel5.Location = '450,40'
    $CopyttextLabel5.size = '200,40'
    $CopyttextLabel5.Text = 'Will have a copy of email sent by users.'
    $CopytSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $CopytSelectedMailBoxNametextLabel6.Location = '350,80'
    $CopytSelectedMailBoxNametextLabel6.Size = '200,40'
    $CopytSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $CopytBox3 = New-Object System.Windows.Forms.GroupBox
    $CopytBox3.Location = '40,340'
    $CopytBox3.size = '700,30'
    $CopytBox3.text = '3. Click Ok to confirm or Cancel:'
    $CopytBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $Copyform.Close()
            $Copyform.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $Copyform.Controls.AddRange(@($CopytBox1, $CopytBox2, $CopytBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $CopytBox1.Controls.AddRange(@($CopyttextLabel1, $CopyMailboxComboBox1))
    $CopytBox2.Controls.AddRange(@($CopyttextLabel3, $CopytSelectedMailboxtextLabel4, $CopyttextLabel5, $CopytSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $Copyform.AcceptButton = $OKButton
    $Copyform.CancelButton = $CancelButton
    #### Activate the form ###
    $Copyform.Add_Shown( { $Copyform.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $Copyform.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($CopytSelectedMailboxtextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Mailbox !!!!!  Trying to enter blank fields is never a good idea.", 'Sent items Copy', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $Copyform.Close()
            $Copyform.Dispose()
            break
        }
        Else {
            $MailBoxPrimarySMTPAddress = get-mailbox $($CopyMailboxComboBox1.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress
            set-mailbox $MailBoxPrimarySMTPAddress -MessageCopyForSendOnBehalfEnabled $True
            set-mailbox $MailBoxPrimarySMTPAddress -MessageCopyForSentAsEnabled $True
            [System.Windows.Forms.MessageBox]::Show("The mailbox ( $($CopyMailboxComboBox1.SelectedItem.ToString()) ) has been set to add a copy of sent emails to the shared mailboxes sent folder.", 'Sent items Copy', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $Copyform.Close()
            $Copyform.Dispose()
            Return SharedMailboxManagementForm
        }
    }
} 
#########################################################################
###  Completed - Create SubForm  Shared Copy to sent items Sub Form  ####
#########################################################################
#endregion Copy
#
#########################################################################################################
##########      Completed - Create 'Shared Mailbox - Management' Sub Forms              #################
#########################################################################################################
#

#########################################################################################################
##########         Create 'User Mailbox - Out Of Office management' Sub Forms               #############
#########################################################################################################
#
#########################################################################
########   Create SubForm Out Of Office - Check current status   ########
#########################################################################
Function OutOfficeCheckForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $OutOfficeCheckForm = New-Object System.Windows.Forms.Form
    $OutOfficeCheckForm.width = 745
    $OutOfficeCheckForm.height = 475
    $OutOfficeCheckForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeCheckForm.Controlbox = $false
    $OutOfficeCheckForm.Icon = $Icon
    $OutOfficeCheckForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeCheckForm.Text = 'User Mailbox - Out Of Office - Check current status.'
    $OutOfficeCheckForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $OutOfficeCheckBox1 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckBox1.Location = '40,20'
    $OutOfficeCheckBox1.size = '650,125'
    $OutOfficeCheckBox1.text = '1. Select a userName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $OutOfficeChecktextLabel2 = New-Object System.Windows.Forms.Label
    $OutOfficeChecktextLabel2.Location = '20,50'
    $OutOfficeChecktextLabel2.size = '200,40'
    $OutOfficeChecktextLabel2.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $OutOfficeCheckMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $OutOfficeCheckMBNameComboBox2.Location = '275,45'
    $OutOfficeCheckMBNameComboBox2.Size = '350, 350'
    $OutOfficeCheckMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $OutOfficeCheckMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $OutOfficeCheckMBNameComboBox2.Sorted = $true;
    $OutOfficeCheckMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $OutOfficeCheckMBNameComboBox2.DataSource = $UserNameList
    $OutOfficeCheckMBNameComboBox2.Add_SelectedIndexChanged( { $OutOfficeCheckSelectedMailBoxNametextLabel6.Text = "$($OutOfficeCheckMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $OutOfficeCheckBox2 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckBox2.Location = '40,170'
    $OutOfficeCheckBox2.size = '650,125'
    $OutOfficeCheckBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $OutOfficeChecktextLabel3 = New-Object System.Windows.Forms.Label
    $OutOfficeChecktextLabel3.Location = '40,40'
    $OutOfficeChecktextLabel3.size = '400,40'
    $OutOfficeChecktextLabel3.Text = 'Check current status:' 
    $OutOfficeCheckSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $OutOfficeCheckSelectedMailBoxNametextLabel6.Location = '100,80'
    $OutOfficeCheckSelectedMailBoxNametextLabel6.Size = '400,40'
    $OutOfficeCheckSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $OutOfficeCheckBox3 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckBox3.Location = '40,320'
    $OutOfficeCheckBox3.size = '650,30'
    $OutOfficeCheckBox3.text = '3. Click Ok to Check current message & status or Cancel:'
    $OutOfficeCheckBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $OutOfficeCheckForm.Close()
            $OutOfficeCheckForm.Dispose()
            Return OutOfficeManagementForm })
    ### Add all the Form controls ### 
    $OutOfficeCheckForm.Controls.AddRange(@($OutOfficeCheckBox1, $OutOfficeCheckBox2, $OutOfficeCheckBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $OutOfficeCheckBox1.Controls.AddRange(@($OutOfficeChecktextLabel2, $OutOfficeCheckMBNameComboBox2))
    $OutOfficeCheckBox2.Controls.AddRange(@($OutOfficeChecktextLabel3, $OutOfficeChecktextLabel5, $OutOfficeCheckSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $OutOfficeCheckForm.AcceptButton = $OKButton
    $OutOfficeCheckForm.CancelButton = $CancelButton
    #### Activate the form ###
    $OutOfficeCheckForm.Add_Shown( { $OutOfficeCheckForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $OutOfficeCheckForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($OutOfficeCheckSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a UserName !!!!!  Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Check current status.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeCheckForm.Close()
            $OutOfficeCheckForm.Dispose()
            break
        }
        #####################################################################
        ############           Check current OOO status       ###############
        #####################################################################
        $UserSamAccountName = get-mailbox $OutOfficeCheckSelectedMailBoxNametextLabel6.Text | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
        $Status = Get-MailboxAutoReplyConfiguration $UserSamAccountName | Select-Object AutoReplyState  
        If ($Status.AutoReplyState -eq 'Disabled') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($OutOfficeCheckMBNameComboBox2.SelectedItem.ToString()) is currently:                      DISABLED - turned off.", 'User Mailbox - Out Of Office - Check current status.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        ElseIf ($Status.AutoReplyState -eq 'Enabled') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for use $($OutOfficeCheckMBNameComboBox2.SelectedItem.ToString()) is currently:                       ENABLED - turned on.", 'User Mailbox - Out Of Office - Check current status.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        $OutOfficeCheckForm.Close()
        $OutOfficeCheckForm.Dispose()
        Return OutOfficeManagementForm
    }
}
#########################################################################
##   Completed - Create SubForm Out Of Office - Check current status   ##
#########################################################################
#
#########################################################################
########         Create SubForm Out Of Office - Turn On          ########
#########################################################################
Function OutOfficeTurnOnForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $OutOfficeTurnOnForm = New-Object System.Windows.Forms.Form
    $OutOfficeTurnOnForm.width = 745
    $OutOfficeTurnOnForm.height = 475
    $OutOfficeTurnOnForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeTurnOnForm.Controlbox = $false
    $OutOfficeTurnOnForm.Icon = $Icon
    $OutOfficeTurnOnForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeTurnOnForm.Text = 'User Mailbox - Out Of Office - Turn On.'
    $OutOfficeTurnOnForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $OutOfficeTurnOnBox1 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOnBox1.Location = '40,20'
    $OutOfficeTurnOnBox1.size = '650,125'
    $OutOfficeTurnOnBox1.text = '1. Select a userName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $OutOfficeTurnOntextLabel2 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOntextLabel2.Location = '20,50'
    $OutOfficeTurnOntextLabel2.size = '200,40'
    $OutOfficeTurnOntextLabel2.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $OutOfficeTurnOnMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $OutOfficeTurnOnMBNameComboBox2.Location = '275,45'
    $OutOfficeTurnOnMBNameComboBox2.Size = '350, 350'
    $OutOfficeTurnOnMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $OutOfficeTurnOnMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $OutOfficeTurnOnMBNameComboBox2.Sorted = $true;
    $OutOfficeTurnOnMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $OutOfficeTurnOnMBNameComboBox2.DataSource = $UserNameList
    $OutOfficeTurnOnMBNameComboBox2.Add_SelectedIndexChanged( { $OutOfficeTurnOnSelectedMailBoxNametextLabel6.Text = "$($OutOfficeTurnOnMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $OutOfficeTurnOnBox2 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOnBox2.Location = '40,170'
    $OutOfficeTurnOnBox2.size = '650,125'
    $OutOfficeTurnOnBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $OutOfficeTurnOntextLabel3 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOntextLabel3.Location = '40,40'
    $OutOfficeTurnOntextLabel3.size = '400,40'
    $OutOfficeTurnOntextLabel3.Text = 'Out Of Office - Turn On for User:' 
    $OutOfficeTurnOnSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOnSelectedMailBoxNametextLabel6.Location = '100,80'
    $OutOfficeTurnOnSelectedMailBoxNametextLabel6.Size = '400,40'
    $OutOfficeTurnOnSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $OutOfficeTurnOnBox3 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOnBox3.Location = '40,320'
    $OutOfficeTurnOnBox3.size = '650,30'
    $OutOfficeTurnOnBox3.text = '3. Click Ok to Turn On Out Of Office or Cancel:'
    $OutOfficeTurnOnBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $OutOfficeTurnOnForm.Close()
            $OutOfficeTurnOnForm.Dispose()
            Return OutOfficeManagementForm })
    ### Add all the Form controls ### 
    $OutOfficeTurnOnForm.Controls.AddRange(@($OutOfficeTurnOnBox1, $OutOfficeTurnOnBox2, $OutOfficeTurnOnBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $OutOfficeTurnOnBox1.Controls.AddRange(@($OutOfficeTurnOntextLabel2, $OutOfficeTurnOnMBNameComboBox2))
    $OutOfficeTurnOnBox2.Controls.AddRange(@($OutOfficeTurnOntextLabel3, $OutOfficeTurnOntextLabel5, $OutOfficeTurnOnSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $OutOfficeTurnOnForm.AcceptButton = $OKButton
    $OutOfficeTurnOnForm.CancelButton = $CancelButton
    #### Activate the form ###
    $OutOfficeTurnOnForm.Add_Shown( { $OutOfficeTurnOnForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $OutOfficeTurnOnForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($OutOfficeTurnOnSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a UserName !!!!!  Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Turn On.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOnForm.Close()
            $OutOfficeTurnOnForm.Dispose()
            break
        }
        #####################################################################
        ############           Check current OOO status       ###############
        #####################################################################
        $UserSamAccountName = get-mailbox $OutOfficeTurnOnSelectedMailBoxNametextLabel6.Text | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
        $Status = Get-MailboxAutoReplyConfiguration $UserSamAccountName | Select-Object AutoReplyState  
        If ($Status.AutoReplyState -eq 'Enabled') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($OutOfficeTurnOnMBNameComboBox2.SelectedItem.ToString())         is already Enabled and turned on.", 'User Mailbox - Out Of Office - Turn On.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOnForm.Close()
            $OutOfficeTurnOnForm.Dispose()
            Return OutOfficeManagementForm
        }
        Else {
            #####################################################################
            #    CHECK - continue if only 1 username in pipe if not exit       ##
            #####################################################################
            if (($UserSamAccountName | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            #               Turn On Out of Office                              ##
            #####################################################################
            Set-MailboxAutoReplyConfiguration $UserSamAccountName –AutoreplyState Enabled
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($OutOfficeTurnOnMBNameComboBox2.SelectedItem.ToString()) has been        ENABLED - turned on.", 'User Mailbox - Out Of Office - Turn On.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOnForm.Close()
            $OutOfficeTurnOnForm.Dispose()
            Return OutOfficeManagementForm
        }
    }
}    
#########################################################################
########    Completed - Create SubForm Out Of Office - Turn On    #######
#########################################################################
#
#########################################################################
########         Create SubForm Out Of Office - Turn Off         ########
#########################################################################
Function OutOfficeTurnOffForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $OutOfficeTurnOffForm = New-Object System.Windows.Forms.Form
    $OutOfficeTurnOffForm.width = 745
    $OutOfficeTurnOffForm.height = 475
    $OutOfficeTurnOffForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeTurnOffForm.Controlbox = $false
    $OutOfficeTurnOffForm.Icon = $Icon
    $OutOfficeTurnOffForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeTurnOffForm.Text = 'User Mailbox - Out Of Office - Turn Off.'
    $OutOfficeTurnOffForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $OutOfficeTurnOffBox1 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOffBox1.Location = '40,20'
    $OutOfficeTurnOffBox1.size = '650,125'
    $OutOfficeTurnOffBox1.text = '1. Select a userName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $OutOfficeTurnOfftextLabel2 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOfftextLabel2.Location = '20,50'
    $OutOfficeTurnOfftextLabel2.size = '200,40'
    $OutOfficeTurnOfftextLabel2.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $OutOfficeTurnOffMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $OutOfficeTurnOffMBNameComboBox2.Location = '275,45'
    $OutOfficeTurnOffMBNameComboBox2.Size = '350, 350'
    $OutOfficeTurnOffMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $OutOfficeTurnOffMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $OutOfficeTurnOffMBNameComboBox2.Sorted = $true;
    $OutOfficeTurnOffMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $OutOfficeTurnOffMBNameComboBox2.DataSource = $UserNameList
    $OutOfficeTurnOffMBNameComboBox2.Add_SelectedIndexChanged( { $OutOfficeTurnOffSelectedMailBoxNametextLabel6.Text = "$($OutOfficeTurnOffMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $OutOfficeTurnOffBox2 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOffBox2.Location = '40,170'
    $OutOfficeTurnOffBox2.size = '650,125'
    $OutOfficeTurnOffBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $OutOfficeTurnOfftextLabel3 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOfftextLabel3.Location = '40,40'
    $OutOfficeTurnOfftextLabel3.size = '400,40'
    $OutOfficeTurnOfftextLabel3.Text = 'Out Of Office - Turn off for user:' 
    $OutOfficeTurnOffSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $OutOfficeTurnOffSelectedMailBoxNametextLabel6.Location = '100,80'
    $OutOfficeTurnOffSelectedMailBoxNametextLabel6.Size = '400,40'
    $OutOfficeTurnOffSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $OutOfficeTurnOffBox3 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeTurnOffBox3.Location = '40,320'
    $OutOfficeTurnOffBox3.size = '650,30'
    $OutOfficeTurnOffBox3.text = '3. Click Ok to Turn Off Out Of Office or Cancel:'
    $OutOfficeTurnOffBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $OutOfficeTurnOffForm.Close()
            $OutOfficeTurnOffForm.Dispose()
            Return OutOfficeManagementForm })
    ### Add all the Form controls ### 
    $OutOfficeTurnOffForm.Controls.AddRange(@($OutOfficeTurnOffBox1, $OutOfficeTurnOffBox2, $OutOfficeTurnOffBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $OutOfficeTurnOffBox1.Controls.AddRange(@($OutOfficeTurnOfftextLabel2, $OutOfficeTurnOffMBNameComboBox2))
    $OutOfficeTurnOffBox2.Controls.AddRange(@($OutOfficeTurnOfftextLabel3, $OutOfficeTurnOfftextLabel5, $OutOfficeTurnOffSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $OutOfficeTurnOffForm.AcceptButton = $OKButton
    $OutOfficeTurnOffForm.CancelButton = $CancelButton
    #### Activate the form ###
    $OutOfficeTurnOffForm.Add_Shown( { $OutOfficeTurnOffForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $OutOfficeTurnOffForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($OutOfficeTurnOffSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a UserName !!!!!  Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Turn Off.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOffForm.Close()
            $OutOfficeTurnOffForm.Dispose()
            break
        }
        #####################################################################
        ############           Check current OOO status       ###############
        #####################################################################
        $UserSamAccountName = get-mailbox $OutOfficeTurnOffSelectedMailBoxNametextLabel6.Text | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
        $Status = Get-MailboxAutoReplyConfiguration $UserSamAccountName | Select-Object AutoReplyState  
        If ($Status.AutoReplyState -eq 'Disabled') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($OutOfficeTurnOffMBNameComboBox2.SelectedItem.ToString())      is already Disabled and turned off.", 'User Mailbox - Out Of Office - Turn Off', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOffForm.Close()
            $OutOfficeTurnOffForm.Dispose()
            Return OutOfficeManagementForm
        }
        Else {
            #####################################################################
            #    CHECK - continue if only 1 username in pipe if not exit        #
            #####################################################################
            if (($UserSamAccountName | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            #               Turn Off Out of Office                              #
            #####################################################################
            Set-MailboxAutoReplyConfiguration $UserSamAccountName –AutoreplyState Disabled
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($OutOfficeTurnOffMBNameComboBox2.SelectedItem.ToString()) has been     DISABLED - turned off.", 'User Mailbox - Out Of Office - Turn Off.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeTurnOffForm.Close()
            $OutOfficeTurnOffForm.Dispose()
            Return OutOfficeManagementForm
        }
    }
}
#########################################################################
########    Completed - Create SubForm Out Of Office - Turn Off   #######
#########################################################################
#
#########################################################################
######  Create SubForm Out Of Office - Add message and turn On   ########
#########################################################################
Function OutOfficeAddTurnOnForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $OutOfficeAddUserForm = New-Object System.Windows.Forms.Form
    $OutOfficeAddUserForm.width = 780
    $OutOfficeAddUserForm.height = 500
    $OutOfficeAddUserForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeAddUserForm.Controlbox = $false
    $OutOfficeAddUserForm.Icon = $Icon
    $OutOfficeAddUserForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeAddUserForm.Text = 'User Mailbox - Out Of Office - Add message and turn On.'
    $OutOfficeAddUserForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $OutOfficeAddUserBox1 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeAddUserBox1.Location = '40,40'
    $OutOfficeAddUserBox1.size = '700,100'
    $OutOfficeAddUserBox1.text = '1. Select a UserName from the dropdown lists & copy the message:'
    ### Create group 1 box text labels. ###
    $OutOfficeAddUsertextLabel1 = New-Object System.Windows.Forms.Label
    $OutOfficeAddUsertextLabel1.Location = '20,40'
    $OutOfficeAddUsertextLabel1.size = '150,40'
    $OutOfficeAddUsertextLabel1.Text = 'UserName:' 
    ### Create group 1 box combo box. ###
    $OutOfficeAddUserUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $OutOfficeAddUserUserNameComboBox1.Location = '325,35'
    $OutOfficeAddUserUserNameComboBox1.Size = '350, 310'
    $OutOfficeAddUserUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $OutOfficeAddUserUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $OutOfficeAddUserUserNameComboBox1.Sorted = $true;
    $OutOfficeAddUserUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $OutOfficeAddUserUserNameComboBox1.DataSource = $UsernameList
    $OutOfficeAddUserUserNameComboBox1.add_SelectedIndexChanged( { $OutOfficeAddUserSelectedUserNametextLabel4.Text = "$($OutOfficeAddUserUserNameComboBox1.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $OutOfficeAddUserBox2 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeAddUserBox2.Location = '40,165'
    $OutOfficeAddUserBox2.size = '700,170'
    $OutOfficeAddUserBox2.text = '2. Type or paste message in box:'
    # Create group 2 box text labels.
    $OutOfficeAddUsertextLabel3 = New-Object System.Windows.Forms.Label
    $OutOfficeAddUsertextLabel3.Location = '40,40'
    $OutOfficeAddUsertextLabel3.size = '100,40'
    $OutOfficeAddUsertextLabel3.Text = 'The User:' 
    $OutOfficeAddUserSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $OutOfficeAddUserSelectedUserNametextLabel4.Location = '30,80'
    $OutOfficeAddUserSelectedUserNametextLabel4.Size = '200,40'
    $OutOfficeAddUserSelectedUserNametextLabel4.ForeColor = 'Blue'
    $OutOfficeAddUsertextLabel5 = New-Object System.Windows.Forms.Label
    $OutOfficeAddUsertextLabel5.Location = '275,40'
    $OutOfficeAddUsertextLabel5.size = '400,30'
    $OutOfficeAddUsertextLabel5.Text = 'Will have the message below added to their Out of Office:'
    # Create group 2 box text box.
    $OutOfficeMessagetextBox2 = New-Object System.Windows.Forms.TextBox
    $OutOfficeMessagetextBox2.Location = '250,80'
    $OutOfficeMessagetextBox2.Size = '450,75'
    $OutOfficeMessagetextBox2.Multiline = $true
    $OutOfficeMessagetextBox2.AcceptsReturn = $true
    $OutOfficeMessagetextBox2.WordWrap = $true
    $OutOfficeMessagetextBox2.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    #$OutOfficeMessagetextBox2.ScrollBars = ScrollBars.Vertical
    ### Create group 3 box in form. ###
    $OutOfficeAddUserBox3 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeAddUserBox3.Location = '40,360'
    $OutOfficeAddUserBox3.size = '700,30'
    $OutOfficeAddUserBox3.text = '3. Click Ok to Add message and turn Out of Office On or Cancel:'
    $OutOfficeAddUserBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,410'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,410'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $OutOfficeAddUserForm.Close()
            $OutOfficeAddUserForm.Dispose()
            Return OutOfficeManagementForm })
    ### Add all the Form controls ### 
    $OutOfficeAddUserForm.Controls.AddRange(@($OutOfficeAddUserBox1, $OutOfficeAddUserBox2, $OutOfficeAddUserBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $OutOfficeAddUserBox1.Controls.AddRange(@($OutOfficeAddUsertextLabel1, $OutOfficeAddUserUserNameComboBox1))
    $OutOfficeAddUserBox2.Controls.AddRange(@($OutOfficeAddUsertextLabel3, $OutOfficeAddUserSelectedUserNametextLabel4, $OutOfficeAddUsertextLabel5, $OutOfficeMessagetextBox2))
    #### Assign the Accept and Cancel options in the form ### 
    $OutOfficeAddUserForm.AcceptButton = $OKButton
    $OutOfficeAddUserForm.CancelButton = $CancelButton
    #### Activate the form ###
    $OutOfficeAddUserForm.Add_Shown( { $OutOfficeAddUserForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $OutOfficeAddUserForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or message     ################ 
        #####################################################################
        if ($OutOfficeAddUserSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!! Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Add message and turn On.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeAddUserForm.Close()
            $OutOfficeAddUserForm.Dispose()
            break
        }
        Elseif ($OutOfficeMessagetextBox2.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to enter a message !!!!! Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Add message and turn On.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeAddUserForm.Close()
            $OutOfficeAddUserForm.Dispose()
            break
        }
        #####################################################################
        ######          Add message and turn On               ###############
        #####################################################################
        $UserSamAccountName = get-mailbox $($OutOfficeAddUserUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        #####################################################################
        #    CHECK - continue if only 1 username in pipe if not exit        #
        #####################################################################
        if (($UserSamAccountName | Measure-Object).count -ne 1) { Mainform }
        #####################################################################
        #            Add message & Turn On Out of Office                    #
        #####################################################################
        Set-MailboxAutoReplyConfiguration -identity $UserSamAccountName -AutoReplyState Enabled -InternalMessage "$($OutOfficeMessagetextBox2.Text)" -ExternalMessage "$($OutOfficeMessagetextBox2.Text)"
        #####################################################################
        ##################  Message complete message  #######################
        #####################################################################
        Add-Type -AssemblyName System.Windows.Forms 
        [System.Windows.Forms.MessageBox]::Show("The message: $($OutOfficeMessagetextBox2.Text) has been added to the Out of Office for $($OutOfficeAddUserUserNameComboBox1.SelectedItem.ToString()) and has been ENABLED - turned on.", "User Mailbox - Out Of Office - Add message and turn On.", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        #####################################################################
        #############   Send email to helpdesk    ###########################
        #####################################################################
        Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($OutOfficeAddUserUserNameComboBox1.SelectedItem.ToString()) has had an Out of Office message added and Turned On." -Body "$($OutOfficeMessagetextBox2.Text)." -SmtpServer mail.scotcourts.local
        $OutOfficeAddUserForm.Close()
        $OutOfficeAddUserForm.Dispose()
        Return OutOfficeManagementForm
    }
} 
#########################################################################
## Completed - Create SubForm Out Of Office - Add message and turn On  ##
#########################################################################
#
#########################################################################
########         Create SubForm Out Of Office - Check Message    ########
#########################################################################
Function OutOfficeCheckMessageForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $OutOfficeCheckMessageForm = New-Object System.Windows.Forms.Form
    $OutOfficeCheckMessageForm.width = 745
    $OutOfficeCheckMessageForm.height = 475
    $OutOfficeCheckMessageForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeCheckMessageForm.Controlbox = $false
    $OutOfficeCheckMessageForm.Icon = $Icon
    $OutOfficeCheckMessageForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeCheckMessageForm.Text = 'User Mailbox - Out Of Office - Check Message.'
    $OutOfficeCheckMessageForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $OutOfficeCheckMessageBox1 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckMessageBox1.Location = '40,20'
    $OutOfficeCheckMessageBox1.size = '650,125'
    $OutOfficeCheckMessageBox1.text = '1. Select a userName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $OutOfficeCheckMessagetextLabel2 = New-Object System.Windows.Forms.Label
    $OutOfficeCheckMessagetextLabel2.Location = '20,50'
    $OutOfficeCheckMessagetextLabel2.size = '200,40'
    $OutOfficeCheckMessagetextLabel2.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $OutOfficeCheckMessageMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $OutOfficeCheckMessageMBNameComboBox2.Location = '275,45'
    $OutOfficeCheckMessageMBNameComboBox2.Size = '350, 350'
    $OutOfficeCheckMessageMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $OutOfficeCheckMessageMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $OutOfficeCheckMessageMBNameComboBox2.Sorted = $true;
    $OutOfficeCheckMessageMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $OutOfficeCheckMessageMBNameComboBox2.DataSource = $UserNameList
    $OutOfficeCheckMessageMBNameComboBox2.Add_SelectedIndexChanged( { $OutOfficeCheckMessageSelectedMailBoxNametextLabel6.Text = "$($OutOfficeCheckMessageMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $OutOfficeCheckMessageBox2 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckMessageBox2.Location = '40,170'
    $OutOfficeCheckMessageBox2.size = '650,125'
    $OutOfficeCheckMessageBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $OutOfficeCheckMessagetextLabel3 = New-Object System.Windows.Forms.Label
    $OutOfficeCheckMessagetextLabel3.Location = '40,40'
    $OutOfficeCheckMessagetextLabel3.size = '400,40'
    $OutOfficeCheckMessagetextLabel3.Text = 'Check current message for User:' 
    $OutOfficeCheckMessageSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $OutOfficeCheckMessageSelectedMailBoxNametextLabel6.Location = '100,80'
    $OutOfficeCheckMessageSelectedMailBoxNametextLabel6.Size = '400,40'
    $OutOfficeCheckMessageSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $OutOfficeCheckMessageBox3 = New-Object System.Windows.Forms.GroupBox
    $OutOfficeCheckMessageBox3.Location = '40,320'
    $OutOfficeCheckMessageBox3.size = '650,30'
    $OutOfficeCheckMessageBox3.text = '3. Click Ok to Check Message or Cancel:'
    $OutOfficeCheckMessageBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $OutOfficeCheckMessageForm.Close()
            $OutOfficeCheckMessageForm.Dispose()
            Return OutOfficeManagementForm })
    ### Add all the Form controls ### 
    $OutOfficeCheckMessageForm.Controls.AddRange(@($OutOfficeCheckMessageBox1, $OutOfficeCheckMessageBox2, $OutOfficeCheckMessageBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $OutOfficeCheckMessageBox1.Controls.AddRange(@($OutOfficeCheckMessagetextLabel2, $OutOfficeCheckMessageMBNameComboBox2))
    $OutOfficeCheckMessageBox2.Controls.AddRange(@($OutOfficeCheckMessagetextLabel3, $OutOfficeCheckMessagetextLabel5, $OutOfficeCheckMessageSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $OutOfficeCheckMessageForm.AcceptButton = $OKButton
    $OutOfficeCheckMessageForm.CancelButton = $CancelButton
    #### Activate the form ###
    $OutOfficeCheckMessageForm.Add_Shown( { $OutOfficeCheckMessageForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $OutOfficeCheckMessageForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($OutOfficeCheckMessageSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a UserName !!!!! Trying to enter blank fields is never a good idea.", 'User Mailbox - Out Of Office - Check Message.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $OutOfficeCheckMessageForm.Close()
            $OutOfficeCheckMessageForm.Dispose()
            break
        }
        #####################################################################
        ############           Check current OOO status       ###############
        #####################################################################
        $UserSamAccountName = get-mailbox $OutOfficeCheckMessageSelectedMailBoxNametextLabel6.Text | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
        Get-MailboxAutoReplyConfiguration -identity $UserSamAccountName | Select-Object ExternalMessage | Out-GridView -Title "The current Out of Office message for $($OutOfficeCheckMessageMBNameComboBox2.SelectedItem.ToString())" -Wait 
        $OutOfficeCheckMessageForm.Close()
        $OutOfficeCheckMessageForm.Dispose()
        Return OutOfficeManagementForm
    }
}
#########################################################################
######  Completed - Create SubForm Out Of Office - Check Message  #######
#########################################################################
#
#########################################################################################################
######        Completed - Create 'User Mailbox - Out Of Office management' Sub Forms        #############
#########################################################################################################
#
#########################################################################################################
###############       Create 'Shared Calendar - User Access Management' Sub Forms        #################
#########################################################################################################
#
######################################################################
####  Create SubForm 'Add - Owner permissions for a User'.   ###
######################################################################
Function AddCalOwnerPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddCalOwnerForm = New-Object System.Windows.Forms.Form
    $AddCalOwnerForm.width = 780
    $AddCalOwnerForm.height = 500
    $AddCalOwnerForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddCalOwnerForm.Controlbox = $false
    $AddCalOwnerForm.Icon = $Icon
    $AddCalOwnerForm.FormBorderStyle = 'Fixed3D'
    $AddCalOwnerForm.Text = 'Calendar - Add Owner (Full Access) Permissions for a User.'
    $AddCalOwnerForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddCalOwnerBox1 = New-Object System.Windows.Forms.GroupBox
    $AddCalOwnerBox1.Location = '40,40'
    $AddCalOwnerBox1.size = '700,125'
    $AddCalOwnerBox1.text = '1. Select a UserName and Calendar from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $AddCalOwnertextLabel1 = New-Object System.Windows.Forms.Label
    $AddCalOwnertextLabel1.Location = '20,40'
    $AddCalOwnertextLabel1.size = '150,40'
    $AddCalOwnertextLabel1.Text = 'UserName:' 
    $AddCalOwnertextLabel2 = New-Object System.Windows.Forms.Label
    $AddCalOwnertextLabel2.Location = '20,80'
    $AddCalOwnertextLabel2.size = '150,40'
    $AddCalOwnertextLabel2.Text = 'Calendar:' 
    ### Create group 1 box combo boxes. ###
    $AddCalOwnerUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $AddCalOwnerUserNameComboBox1.Location = '325,35'
    $AddCalOwnerUserNameComboBox1.Size = '350, 310'
    $AddCalOwnerUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $AddCalOwnerUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $AddCalOwnerUserNameComboBox1.Sorted = $true;
    $AddCalOwnerUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddCalOwnerUserNameComboBox1.DataSource = $UsernameList
    $AddCalOwnerUserNameComboBox1.add_SelectedIndexChanged( { $AddCalOwnerSelectedUserNametextLabel4.Text = "$($AddCalOwnerUserNameComboBox1.SelectedItem.ToString())" })
    $AddCalOwnerMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddCalOwnerMBNameComboBox2.Location = '325,75'
    $AddCalOwnerMBNameComboBox2.Size = '350, 350'
    $AddCalOwnerMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddCalOwnerMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddCalOwnerMBNameComboBox2.Sorted = $true;
    $AddCalOwnerMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddCalOwnerMBNameComboBox2.DataSource = $Sharedmailboxes 
    $AddCalOwnerMBNameComboBox2.add_SelectedIndexChanged( { $AddCalOwnerSelectedMailBoxNametextLabel6.Text = "$($AddCalOwnerMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddCalOwnerBox2 = New-Object System.Windows.Forms.GroupBox
    $AddCalOwnerBox2.Location = '40,190'
    $AddCalOwnerBox2.size = '700,125'
    $AddCalOwnerBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $AddCalOwnertextLabel3 = New-Object System.Windows.Forms.Label
    $AddCalOwnertextLabel3.Location = '40,40'
    $AddCalOwnertextLabel3.size = '100,40'
    $AddCalOwnertextLabel3.Text = 'The User:' 
    $AddCalOwnerSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $AddCalOwnerSelectedUserNametextLabel4.Location = '30,80'
    $AddCalOwnerSelectedUserNametextLabel4.Size = '200,40'
    $AddCalOwnerSelectedUserNametextLabel4.ForeColor = 'Blue'
    $AddCalOwnertextLabel5 = New-Object System.Windows.Forms.Label
    $AddCalOwnertextLabel5.Location = '275,40'
    $AddCalOwnertextLabel5.size = '400,40'
    $AddCalOwnertextLabel5.Text = 'Will have Owner permissions added to the Calendar:'
    $AddCalOwnerSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $AddCalOwnerSelectedMailBoxNametextLabel6.Location = '350,80'
    $AddCalOwnerSelectedMailBoxNametextLabel6.Size = '200,40'
    $AddCalOwnerSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $AddCalOwnerBox3 = New-Object System.Windows.Forms.GroupBox
    $AddCalOwnerBox3.Location = '40,340'
    $AddCalOwnerBox3.size = '700,30'
    $AddCalOwnerBox3.text = '3. Click Ok to add Owner permissions or Cancel:'
    $AddCalOwnerBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $AddCalOwnerForm.Close()
            $AddCalOwnerForm.Dispose()
            Return SharedCalendarManagementForm })
    ### Add all the Form controls ### 
    $AddCalOwnerForm.Controls.AddRange(@($AddCalOwnerBox1, $AddCalOwnerBox2, $AddCalOwnerBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $AddCalOwnerBox1.Controls.AddRange(@($AddCalOwnertextLabel1, $AddCalOwnertextLabel2, $AddCalOwnerUserNameComboBox1, $AddCalOwnerMBNameComboBox2))
    $AddCalOwnerBox2.Controls.AddRange(@($AddCalOwnertextLabel3, $AddCalOwnerSelectedUserNametextLabel4, $AddCalOwnertextLabel5, $AddCalOwnerSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $AddCalOwnerForm.AcceptButton = $OKButton
    $AddCalOwnerForm.CancelButton = $CancelButton
    #### Activate the form ###
    $AddCalOwnerForm.Add_Shown( { $AddCalOwnerForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $AddCalOwnerForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($AddCalOwnerSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!! Trying to enter blank fields is never a good idea.", 'Calendar - Add Owner Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalOwnerForm.Close()
            $AddCalOwnerForm.Dispose()
            break
        }
        Elseif ($AddCalOwnerSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Calendar !!!!! Trying to enter blank fields is never a good idea.", 'Calendar - Add Owner Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalOwnerForm.Close()
            $AddCalOwnerForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($AddCalOwnerUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailBoxPrimarySMTPAddress = get-mailbox $($AddCalOwnerMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        #########       Check if User already has owner access  #############
        #####################################################################
        $Status = Get-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user "$UserSamAccountName"
        If ($Status.AccessRights -eq 'Owner') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalOwnerUserNameComboBox1.SelectedItem.ToString()) ) already has Owner Permissions to the ( $($AddCalOwnerMBNameComboBox2.SelectedItem.ToString()) ) calendar.", 'Calendar - Add Owner Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        If ($Status.AccessRights -eq 'Reviewer') {
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalOwnerUserNameComboBox1.SelectedItem.ToString()) ) already has REVIEWER Permissions to the ( $($AddCalOwnerMBNameComboBox2.SelectedItem.ToString()) ) calendar.Remove the REVIEWER permissions before adding Owner permissions", 'Calendar - Add Owner Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalOwnerForm.Close()
            $AddCalOwnerForm.Dispose()
            Return SharedCalendarManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######       Add Owner permission & deleted folder access     #######
            #####################################################################
            Add-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user "$UserSamAccountName" -AccessRights Owner
            Add-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\deleted items" -user "$UserSamAccountName" -AccessRights Contributor
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalOwnerUserNameComboBox1.SelectedItem.ToString()) ) has had Owner Permissions added to the ( $($AddCalOwnerMBNameComboBox2.SelectedItem.ToString()) ) calendar. The user needs to close and re-open Outlook to pick up the permissions.", 'Calendar - Add Owner Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($AddCalOwnerUserNameComboBox1.SelectedItem.ToString()) has had Owner permissions added to the $($AddCalOwnerMBNameComboBox2.SelectedItem.ToString()) calendar" -Body 'The user needs to close and re-open Outlook to pick up the permissions.', -SmtpServer mail.scotcourts.local
            $AddCalOwnerForm.Close()
            $AddCalOwnerForm.Dispose()
            Return SharedCalendarManagementForm
        }
    }
}
#########################################################################
##    Completed -Create SubForm Add - Owner permissions for a User     ##
#########################################################################
#
#########################################################################
#####   Create SubForm 'Add - Reviewer permissions for a User'.     #####
#########################################################################
Function AddCalReviewerPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddCalReviewerForm = New-Object System.Windows.Forms.Form
    $AddCalReviewerForm.width = 780
    $AddCalReviewerForm.height = 500
    $AddCalReviewerForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddCalReviewerForm.Controlbox = $false
    $AddCalReviewerForm.Icon = $Icon
    $AddCalReviewerForm.FormBorderStyle = 'Fixed3D'
    $AddCalReviewerForm.Text = 'Calendar - Add Reviewer Permissions for a User.'
    $AddCalReviewerForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddCalReviewerBox1 = New-Object System.Windows.Forms.GroupBox
    $AddCalReviewerBox1.Location = '40,40'
    $AddCalReviewerBox1.size = '700,125'
    $AddCalReviewerBox1.text = '1. Select a UserName and Calendar from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $AddCalReviewertextLabel1 = New-Object System.Windows.Forms.Label
    $AddCalReviewertextLabel1.Location = '20,40'
    $AddCalReviewertextLabel1.size = '150,40'
    $AddCalReviewertextLabel1.Text = 'UserName:' 
    $AddCalReviewertextLabel2 = New-Object System.Windows.Forms.Label
    $AddCalReviewertextLabel2.Location = '20,80'
    $AddCalReviewertextLabel2.size = '150,40'
    $AddCalReviewertextLabel2.Text = 'Calendar:' 
    ### Create group 1 box combo boxes. ###
    $AddCalReviewerUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $AddCalReviewerUserNameComboBox1.Location = '325,35'
    $AddCalReviewerUserNameComboBox1.Size = '350, 310'
    $AddCalReviewerUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $AddCalReviewerUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $AddCalReviewerUserNameComboBox1.Sorted = $true;
    $AddCalReviewerUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddCalReviewerUserNameComboBox1.DataSource = $UsernameList
    $AddCalReviewerUserNameComboBox1.add_SelectedIndexChanged( { $AddCalReviewerSelectedUserNametextLabel4.Text = "$($AddCalReviewerUserNameComboBox1.SelectedItem.ToString())" })
    $AddCalReviewerMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddCalReviewerMBNameComboBox2.Location = '325,75'
    $AddCalReviewerMBNameComboBox2.Size = '350, 350'
    $AddCalReviewerMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddCalReviewerMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddCalReviewerMBNameComboBox2.Sorted = $true;
    $AddCalReviewerMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddCalReviewerMBNameComboBox2.DataSource = $Sharedmailboxes 
    $AddCalReviewerMBNameComboBox2.add_SelectedIndexChanged( { $AddCalReviewerSelectedMailBoxNametextLabel6.Text = "$($AddCalReviewerMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddCalReviewerBox2 = New-Object System.Windows.Forms.GroupBox
    $AddCalReviewerBox2.Location = '40,190'
    $AddCalReviewerBox2.size = '700,125'
    $AddCalReviewerBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $AddCalReviewertextLabel3 = New-Object System.Windows.Forms.Label
    $AddCalReviewertextLabel3.Location = '40,40'
    $AddCalReviewertextLabel3.size = '100,40'
    $AddCalReviewertextLabel3.Text = 'The User:' 
    $AddCalReviewerSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $AddCalReviewerSelectedUserNametextLabel4.Location = '30,80'
    $AddCalReviewerSelectedUserNametextLabel4.ForeColor = 'Blue'
    $AddCalReviewerSelectedUserNametextLabel4.Size = '200,40'
    $AddCalReviewertextLabel5 = New-Object System.Windows.Forms.Label
    $AddCalReviewertextLabel5.Location = '275,40'
    $AddCalReviewertextLabel5.size = '400,40'
    $AddCalReviewertextLabel5.Text = 'Will have Reviewer permissions added to the Calendar:'
    $AddCalReviewerSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $AddCalReviewerSelectedMailBoxNametextLabel6.Location = '350,80'
    $AddCalReviewerSelectedMailBoxNametextLabel6.Size = '200,40'
    $AddCalReviewerSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $AddCalReviewerBox3 = New-Object System.Windows.Forms.GroupBox
    $AddCalReviewerBox3.Location = '40,340'
    $AddCalReviewerBox3.size = '700,30'
    $AddCalReviewerBox3.text = '3. Click Ok to add Reviewer permissions or Cancel:'
    $AddCalReviewerBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $AddCalReviewerForm.Close()
            $AddCalReviewerForm.Dispose()
            Return SharedCalendarManagementForm })
    ### Add all the Form controls ### 
    $AddCalReviewerForm.Controls.AddRange(@($AddCalReviewerBox1, $AddCalReviewerBox2, $AddCalReviewerBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $AddCalReviewerBox1.Controls.AddRange(@($AddCalReviewertextLabel1, $AddCalReviewertextLabel2, $AddCalReviewerUserNameComboBox1, $AddCalReviewerMBNameComboBox2))
    $AddCalReviewerBox2.Controls.AddRange(@($AddCalReviewertextLabel3, $AddCalReviewerSelectedUserNametextLabel4, $AddCalReviewertextLabel5, $AddCalReviewerSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $AddCalReviewerForm.AcceptButton = $OKButton
    $AddCalReviewerForm.CancelButton = $CancelButton
    #### Activate the form ###
    $AddCalReviewerForm.Add_Shown( { $AddCalReviewerForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $AddCalReviewerForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($AddCalReviewerSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!! Trying to enter blank fields is never a good idea.", 'Calendar - Add Reviewer Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalReviewerForm.Close()
            $AddCalReviewerForm.Dispose()
            break
        }
        Elseif ($AddCalReviewerSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Calendar !!!!! Trying to enter blank fields is never a good idea.", 'Calendar - Add Reviewer Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalReviewerForm.Close()
            $AddCalReviewerForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($AddCalReviewerUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailBoxPrimarySMTPAddress = get-mailbox $($AddCalReviewerMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        #########    Check if User already has reviewer access  #############
        #####################################################################
        $Status = Get-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user "$UserSamAccountName"
        If ($Status.AccessRights -eq 'Reviewer') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalReviewerUserNameComboBox1.SelectedItem.ToString()) ) already has Reviewer Permissions to the ( $($AddCalReviewerMBNameComboBox2.SelectedItem.ToString()) ) calendar.", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        If ($Status.AccessRights -eq 'Owner') {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalReviewerUserNameComboBox1.SelectedItem.ToString()) ) already has OWNER Permissions to the ( $($AddCalReviewerMBNameComboBox2.SelectedItem.ToString()) ) calendar. Remove the OWNER permissions before adding Reviewer permissions", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddCalReviewerForm.Close()
            $AddCalReviewerForm.Dispose()
            Return SharedCalendarManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######    Add Reviewer permission & deleted folder access     #######
            #####################################################################
            Add-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user "$UserSamAccountName" -AccessRights Reviewer
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($AddCalReviewerUserNameComboBox1.SelectedItem.ToString()) ) has had Reviewer Permissions added to the ( $($AddCalReviewerMBNameComboBox2.SelectedItem.ToString()) ) calendar. The user needs to close and re-open Outlook to pick up the permissions.", 'Calendar - Add Reviewer Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($AddCalReviewerUserNameComboBox1.SelectedItem.ToString()) has had Reviewer permissions added to the $($AddCalReviewerMBNameComboBox2.SelectedItem.ToString()) calendar" -Body 'The user needs to close and re-open Outlook to pick up the permissions.' -SmtpServer mail.scotcourts.local
            $AddCalReviewerForm.Close()
            $AddCalReviewerForm.Dispose()
            Return SharedCalendarManagementForm
        }   
    }
}
#########################################################################
##   Completed -Create SubForm Add - Reviewer permissions for a User   ##
#########################################################################
#
#########################################################################
#########       Create SubForm Check Calendar permissions     ###########
#########################################################################
Function CheckCalPermAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $CheckCalPermForm = New-Object System.Windows.Forms.Form
    $CheckCalPermForm.width = 745
    $CheckCalPermForm.height = 475
    $CheckCalPermForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $CheckCalPermForm.Controlbox = $false
    $CheckCalPermForm.Icon = $Icon
    $CheckCalPermForm.FormBorderStyle = 'Fixed3D'
    $CheckCalPermForm.Text = 'Calendar - Check current Calendar Permissions.'
    $CheckCalPermForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $CheckCalPermBox1 = New-Object System.Windows.Forms.GroupBox
    $CheckCalPermBox1.Location = '40,20'
    $CheckCalPermBox1.size = '650,125'
    $CheckCalPermBox1.text = '1. Select a MailBoxName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $CheckCalPermtextLabel2 = New-Object System.Windows.Forms.Label
    $CheckCalPermtextLabel2.Location = '20,50'
    $CheckCalPermtextLabel2.size = '200,40'
    $CheckCalPermtextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $CheckCalPermMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $CheckCalPermMBNameComboBox2.Location = '275,45'
    $CheckCalPermMBNameComboBox2.Size = '350, 350'
    $CheckCalPermMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $CheckCalPermMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $CheckCalPermMBNameComboBox2.Sorted = $true;
    $CheckCalPermMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $CheckCalPermMBNameComboBox2.DataSource = $Sharedmailboxes
    $CheckCalPermMBNameComboBox2.Add_SelectedIndexChanged( { $CheckCalPermSelectedMailBoxNametextLabel6.Text = "$($CheckCalPermMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $CheckCalPermBox2 = New-Object System.Windows.Forms.GroupBox
    $CheckCalPermBox2.Location = '40,170'
    $CheckCalPermBox2.size = '650,125'
    $CheckCalPermBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $CheckCalPermtextLabel3 = New-Object System.Windows.Forms.Label
    $CheckCalPermtextLabel3.Location = '40,40'
    $CheckCalPermtextLabel3.size = '400,40'
    $CheckCalPermtextLabel3.Text = 'Check current Calendar permissions:' 
    $CheckCalPermSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $CheckCalPermSelectedMailBoxNametextLabel6.Location = '100,80'
    $CheckCalPermSelectedMailBoxNametextLabel6.Size = '400,40'
    $CheckCalPermSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $CheckCalPermBox3 = New-Object System.Windows.Forms.GroupBox
    $CheckCalPermBox3.Location = '40,320'
    $CheckCalPermBox3.size = '650,30'
    $CheckCalPermBox3.text = '3. Click Ok to Check current Calendar permissions or Cancel:'
    $CheckCalPermBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $CheckCalPermForm.Close()
            $CheckCalPermForm.Dispose()
            Return SharedCalendarManagementForm })
    ### Add all the Form controls ### 
    $CheckCalPermForm.Controls.AddRange(@($CheckCalPermBox1, $CheckCalPermBox2, $CheckCalPermBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $CheckCalPermBox1.Controls.AddRange(@($CheckCalPermtextLabel2, $CheckCalPermMBNameComboBox2))
    $CheckCalPermBox2.Controls.AddRange(@($CheckCalPermtextLabel3, $CheckCalPermtextLabel5, $CheckCalPermSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $CheckCalPermForm.AcceptButton = $OKButton
    $CheckCalPermForm.CancelButton = $CancelButton
    #### Activate the form ###
    $CheckCalPermForm.Add_Shown( { $CheckCalPermForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckCalPermForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($CheckCalPermSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Calendar !!!!!Trying to enter blank fields is never a good idea.", 'Calendar - Check current Calendar Permissions.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckCalPermForm.Close()
            $CheckCalPermForm.Dispose()
            break
        }
        #####################################################################
        ############         Check Calendar permissions       ###############
        #####################################################################
        $MailboxName = Get-Mailbox -identity $CheckCalPermSelectedMailBoxNametextLabel6.Text | Select-Object primarysmtpaddress 
        #primarysmtpaddress
        Get-MailboxFolderPermission -identity ($MailBoxname.PrimarySmtpAddress + ':\calendar') | Select-Object User, AccessRights | Sort-Object user | Out-GridView -Title "List of Users with permissions on $($CheckCalPermMBNameComboBox2.SelectedItem.ToString()) calendar" -Wait 
        $CheckCalPermForm.Close()
        $CheckCalPermForm.Dispose()
        Return SharedCalendarManagementForm
    }
}
#########################################################################
######   Completed - Create SubForm Check Calendar permissions ##########
#########################################################################
#
#########################################################################
######          Create SubForm Remove Calendar  access          #########
#########################################################################
Function RemoveCalendarPermAccessPermissionForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $RemoveCalendarPermForm = New-Object System.Windows.Forms.Form
    $RemoveCalendarPermForm.width = 745
    $RemoveCalendarPermForm.height = 475
    $RemoveCalendarPermForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RemoveCalendarPermForm.Controlbox = $false
    $RemoveCalendarPermForm.Icon = $Icon
    $RemoveCalendarPermForm.FormBorderStyle = 'Fixed3D'
    $RemoveCalendarPermForm.Text = 'Calendar - Remove Calendar Access Permissions for a User.'
    $RemoveCalendarPermForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $RemoveCalendarPermBox1 = New-Object System.Windows.Forms.GroupBox
    $RemoveCalendarPermBox1.Location = '40,20'
    $RemoveCalendarPermBox1.size = '650,125'
    $RemoveCalendarPermBox1.text = '1. Select a UserName and Calendar from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $RemoveCalendarPermtextLabel1 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermtextLabel1.Location = '20,40'
    $RemoveCalendarPermtextLabel1.size = '100,40'
    $RemoveCalendarPermtextLabel1.Text = 'UserName:' 
    $RemoveCalendarPermtextLabel2 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermtextLabel2.Location = '20,80'
    $RemoveCalendarPermtextLabel2.size = '100,40'
    $RemoveCalendarPermtextLabel2.Text = 'Calendar:' 
    ### Create group 1 box combo boxes. ###
    $RemoveCalendarPermUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $RemoveCalendarPermUserNameComboBox1.Location = '275,35'
    $RemoveCalendarPermUserNameComboBox1.Size = '350, 310'
    $RemoveCalendarPermUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $RemoveCalendarPermUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $RemoveCalendarPermUserNameComboBox1.Sorted = $true;
    $RemoveCalendarPermUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveCalendarPermUserNameComboBox1.DataSource = $UsernameList
    $RemoveCalendarPermUserNameComboBox1.Add_SelectedIndexChanged( { $RemoveCalendarPermSelectedUserNametextLabel4.Text = "$($RemoveCalendarPermUserNameComboBox1.SelectedItem.ToString())" })
    $RemoveCalendarPermMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $RemoveCalendarPermMBNameComboBox2.Location = '275,75'
    $RemoveCalendarPermMBNameComboBox2.Size = '350, 350'
    $RemoveCalendarPermMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $RemoveCalendarPermMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $RemoveCalendarPermMBNameComboBox2.Sorted = $true;
    $RemoveCalendarPermMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RemoveCalendarPermMBNameComboBox2.DataSource = $Sharedmailboxes 
    $RemoveCalendarPermMBNameComboBox2.Add_SelectedIndexChanged( { $RemoveCalendarPermSelectedMailBoxNametextLabel6.Text = "$($RemoveCalendarPermMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $RemoveCalendarPermBox2 = New-Object System.Windows.Forms.GroupBox
    $RemoveCalendarPermBox2.Location = '40,170'
    $RemoveCalendarPermBox2.size = '650,125'
    $RemoveCalendarPermBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $RemoveCalendarPermtextLabel3 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermtextLabel3.Location = '40,40'
    $RemoveCalendarPermtextLabel3.size = '100,40'
    $RemoveCalendarPermtextLabel3.Text = 'The User:' 
    $RemoveCalendarPermSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermSelectedUserNametextLabel4.Location = '30,80'
    $RemoveCalendarPermSelectedUserNametextLabel4.Size = '200,40'
    $RemoveCalendarPermSelectedUserNametextLabel4.ForeColor = 'Blue'
    $RemoveCalendarPermtextLabel5 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermtextLabel5.Location = '175,40'
    $RemoveCalendarPermtextLabel5.size = '450,40'
    $RemoveCalendarPermtextLabel5.Text = 'Will have Calendar Access permissions removed from:'
    $RemoveCalendarPermSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $RemoveCalendarPermSelectedMailBoxNametextLabel6.Location = '350,80'
    $RemoveCalendarPermSelectedMailBoxNametextLabel6.Size = '200,40'
    $RemoveCalendarPermSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $RemoveCalendarPermBox3 = New-Object System.Windows.Forms.GroupBox
    $RemoveCalendarPermBox3.Location = '40,320'
    $RemoveCalendarPermBox3.size = '650,30'
    $RemoveCalendarPermBox3.text = '3. Click Continue to remove Calendar Access permissions or Cancel:'
    $RemoveCalendarPermBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $RemoveCalendarPermForm.Close()
            $RemoveCalendarPermForm.Dispose()
            Return SharedCalendarManagementForm })
    ### Add all the Form controls ### 
    $RemoveCalendarPermForm.Controls.AddRange(@($RemoveCalendarPermBox1, $RemoveCalendarPermBox2, $RemoveCalendarPermBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $RemoveCalendarPermBox1.Controls.AddRange(@($RemoveCalendarPermtextLabel1, $RemoveCalendarPermtextLabel2, $RemoveCalendarPermUserNameComboBox1, $RemoveCalendarPermMBNameComboBox2))
    $RemoveCalendarPermBox2.Controls.AddRange(@($RemoveCalendarPermtextLabel3, $RemoveCalendarPermSelectedUserNametextLabel4, $RemoveCalendarPermtextLabel5, $RemoveCalendarPermSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $RemoveCalendarPermForm.AcceptButton = $OKButton
    $RemoveCalendarPermForm.CancelButton = $CancelButton
    #### Activate the form ###
    $RemoveCalendarPermForm.add_Shown( { $RemoveCalendarPermForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $RemoveCalendarPermForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($RemoveCalendarPermSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!Trying to enter blank fields is never a good idea.", 'Calendar - Remove Calendar Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveCalendarPermForm.Close()
            $RemoveCalendarPermForm.Dispose()
            break
        }
        Elseif ($RemoveCalendarPermSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Calendar !!!!!Trying to enter blank fields is never a good idea.", 'Calendar - Remove Calendar Access Permissions for a User.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveCalendarPermForm.Close()
            $RemoveCalendarPermForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get mailbox primary smtpaddress from mailbox display name:  ####
        #####################################################################
        $UserSamAccountName = get-mailbox $($RemoveCalendarPermUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailBoxPrimarySMTPAddress = get-mailbox $($RemoveCalendarPermMBNameComboBox2.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
        #####################################################################
        ########        Check if User already has calendar access  ##########
        #####################################################################
        $Status = Get-MailboxFolderPermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user "$UserSamAccountName"
        If (($Status.AccessRights -ne 'Owner') -and ($Status.AccessRights -ne 'Reviewer')) {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ($($RemoveCalendarPermUserNameComboBox1.SelectedItem.ToString())) does not have Calendar Access Permissions to the ($($RemoveCalendarPermMBNameComboBox2.SelectedItem.ToString())) Calendar.", 'Calendar - Remove Calendar Access Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RemoveCalendarPermForm.Close()
            $RemoveCalendarPermForm.Dispose()
            Return SharedCalendarManagementForm
        }
        Else {
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ######            Remove Full access for user        ################
            #####################################################################
            Remove-mailboxfolderpermission -identity "${MailBoxPrimarySMTPAddress}:\calendar" -user $UserSamAccountName -confirm:$false
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ($($RemoveCalendarPermUserNameComboBox1.SelectedItem.ToString())) has had Calendar Access Permissions removed on the ($($RemoveCalendarPermMBNameComboBox2.SelectedItem.ToString())) Calendar.", 'Calendar - Remove Calendar Access Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($RemoveCalendarPermUserNameComboBox1.SelectedItem.ToString()) has had Calendar permissions removed from the $($RemoveCalendarPermMBNameComboBox2.SelectedItem.ToString()) calendar" -Body 'The user needs to close and re-open Outlook.' -SmtpServer mail.scotcourts.local
            $RemoveCalendarPermForm.Close()
            $RemoveCalendarPermForm.Dispose()
            Return SharedCalendarManagementForm
        }
    }
}
#########################################################################
#####     Completed - Create SubForm Remove Calendar permission  ########
#########################################################################
#
#########################################################################################################
##########     Completed - Create 'Shared Calendar - User Access Management' Sub Forms  #################
#########################################################################################################
#
#########################################################################################################
##########       Create 'Shared Mailbox - Distribution List Management' Sub Forms       #################
#########################################################################################################
#
######################################################################
####   Create SubForm 'Distribution List - Add User to a List     ####
######################################################################
Function DistributionListAddUserForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $DistListAddUserForm = New-Object System.Windows.Forms.Form
    $DistListAddUserForm.width = 780
    $DistListAddUserForm.height = 500
    $DistListAddUserForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DistListAddUserForm.Controlbox = $false
    $DistListAddUserForm.Icon = $Icon
    $DistListAddUserForm.FormBorderStyle = 'Fixed3D'
    $DistListAddUserForm.Text = 'Mailbox - Distribution List - Add User to a List.'
    $DistListAddUserForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $DistListAddUserBox1 = New-Object System.Windows.Forms.GroupBox
    $DistListAddUserBox1.Location = '40,40'
    $DistListAddUserBox1.size = '700,125'
    $DistListAddUserBox1.text = '1. Select a UserName and Distribution list from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $DistListAddUsertextLabel1 = New-Object System.Windows.Forms.Label
    $DistListAddUsertextLabel1.Location = '20,40'
    $DistListAddUsertextLabel1.size = '150,40'
    $DistListAddUsertextLabel1.Text = 'UserName:' 
    $DistListAddUsertextLabel2 = New-Object System.Windows.Forms.Label
    $DistListAddUsertextLabel2.Location = '20,80'
    $DistListAddUsertextLabel2.size = '150,40'
    $DistListAddUsertextLabel2.Text = 'Distribution list:' 
    ### Create group 1 box combo boxes. ###
    $DistListAddUserUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $DistListAddUserUserNameComboBox1.Location = '325,35'
    $DistListAddUserUserNameComboBox1.Size = '350, 310'
    $DistListAddUserUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $DistListAddUserUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $DistListAddUserUserNameComboBox1.Sorted = $true;
    $DistListAddUserUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $DistListAddUserUserNameComboBox1.DataSource = $UsernameList
    $DistListAddUserUserNameComboBox1.add_SelectedIndexChanged( { $DistListAddUserSelectedUserNametextLabel4.Text = "$($DistListAddUserUserNameComboBox1.SelectedItem.ToString())" })
    $DistListAddUserMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $DistListAddUserMBNameComboBox2.Location = '325,75'
    $DistListAddUserMBNameComboBox2.Size = '350, 350'
    $DistListAddUserMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $DistListAddUserMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $DistListAddUserMBNameComboBox2.Sorted = $true;
    $DistListAddUserMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $DistListAddUserMBNameComboBox2.DataSource = $DistributionLists 
    $DistListAddUserMBNameComboBox2.add_SelectedIndexChanged( { $DistListAddUserSelectedMailBoxNametextLabel6.Text = "$($DistListAddUserMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $DistListAddUserBox2 = New-Object System.Windows.Forms.GroupBox
    $DistListAddUserBox2.Location = '40,190'
    $DistListAddUserBox2.size = '700,125'
    $DistListAddUserBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $DistListAddUsertextLabel3 = New-Object System.Windows.Forms.Label
    $DistListAddUsertextLabel3.Location = '40,40'
    $DistListAddUsertextLabel3.size = '100,40'
    $DistListAddUsertextLabel3.Text = 'The User:' 
    $DistListAddUserSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $DistListAddUserSelectedUserNametextLabel4.Location = '30,80'
    $DistListAddUserSelectedUserNametextLabel4.Size = '200,40'
    $DistListAddUserSelectedUserNametextLabel4.ForeColor = 'Blue'
    $DistListAddUsertextLabel5 = New-Object System.Windows.Forms.Label
    $DistListAddUsertextLabel5.Location = '275,40'
    $DistListAddUsertextLabel5.size = '400,40'
    $DistListAddUsertextLabel5.Text = 'Will be added to the Distribution list:'
    $DistListAddUserSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $DistListAddUserSelectedMailBoxNametextLabel6.Location = '350,80'
    $DistListAddUserSelectedMailBoxNametextLabel6.Size = '200,40'
    $DistListAddUserSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $DistListAddUserBox3 = New-Object System.Windows.Forms.GroupBox
    $DistListAddUserBox3.Location = '40,340'
    $DistListAddUserBox3.size = '700,30'
    $DistListAddUserBox3.text = '3. Click Ok to add User to list or Cancel:'
    $DistListAddUserBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $DistListAddUserForm.Close()
            $DistListAddUserForm.Dispose()
            Return DistributionListManagementForm })
    ### Add all the Form controls ### 
    $DistListAddUserForm.Controls.AddRange(@($DistListAddUserBox1, $DistListAddUserBox2, $DistListAddUserBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $DistListAddUserBox1.Controls.AddRange(@($DistListAddUsertextLabel1, $DistListAddUsertextLabel2, $DistListAddUserUserNameComboBox1, $DistListAddUserMBNameComboBox2))
    $DistListAddUserBox2.Controls.AddRange(@($DistListAddUsertextLabel3, $DistListAddUserSelectedUserNametextLabel4, $DistListAddUsertextLabel5, $DistListAddUserSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $DistListAddUserForm.AcceptButton = $OKButton
    $DistListAddUserForm.CancelButton = $CancelButton
    #### Activate the form ###
    $DistListAddUserForm.Add_Shown( { $DistListAddUserForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $DistListAddUserForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($DistListAddUserSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!Trying to enter blank fields is never a good idea.", 'Mailbox - Distribution List - Add User to a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListAddUserForm.Close()
            $DistListAddUserForm.Dispose()
            break
        }
        Elseif ($DistListAddUserSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Distribution List !!!!!Trying to enter blank fields is never a good idea.", 'Mailbox - Distribution List - Add User to a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListAddUserForm.Close()
            $DistListAddUserForm.Dispose()
            break
        }
        #####################################################################
        #########                get SamAccountNames            #############
        #####################################################################
        $UserSamAccountName = get-mailbox $($DistListAddUserSelectedUserNametextLabel4.Text) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailboxSamAccountname = Get-DistributionGroup -Filter "Name -eq '$($DistListAddUserSelectedMailBoxNametextLabel6.Text)'" | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName 
        #####################################################################
        #########        Check if User already is a member      #############
        #####################################################################
        $DistributionGroupMembers = Get-DistributionGroupMember $MailboxSamAccountname | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
    
        Write-Output "User $UserSamAccountName"
        Write-Output "Mailbox $MailboxSamAccountName"
        Write-Output "List $DistributionGroupMembers"
   
        If ($DistributionGroupMembers -notcontains $UserSamAccountName) {
            Add-DistributionGroupMember -Identity $MailboxSamAccountname -Member $UserSamAccountName
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user $($DistListAddUserUserNameComboBox1.SelectedItem.ToString()) has been added to Distribution list $($DistListAddUserMBNameComboBox2.SelectedItem.ToString())`n`nThe user needs to close and re-open Outlook.", 'Mailbox - Distribution List - Add User to a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListAddUserForm.Close()
            $DistListAddUserForm.Dispose()
            Return DistributionListManagementForm
        }
        Else {
            { Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user $($DistListAddUserUserNameComboBox1.SelectedItem.ToString()) is already a member of the $($DistListAddUserMBNameComboBox2.SelectedItem.ToString()) distribution list.", 'Mailbox - Distribution List - Add User to a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $DistListAddUserForm.Close()
                $DistListAddUserForm.Dispose()
                Return DistributionListManagementForm }
        }
    }
}
#########################################################################
#####   Completed -  Create SubForm Dist List - Add user to a list  #####
#########################################################################
#
#########################################################################
###   Create SubForm 'Distribution List - Remove User from a List    ####
#########################################################################
Function DistributionListRemoveUsererForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $DistListRemoveUserForm = New-Object System.Windows.Forms.Form
    $DistListRemoveUserForm.width = 780
    $DistListRemoveUserForm.height = 500
    $DistListRemoveUserForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DistListRemoveUserForm.Controlbox = $false
    $DistListRemoveUserForm.Icon = $Icon
    $DistListRemoveUserForm.FormBorderStyle = 'Fixed3D'
    $DistListRemoveUserForm.Text = 'Mailbox - Distribution List - Remove User from a List.'
    $DistListRemoveUserForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $DistListRemoveUserBox1 = New-Object System.Windows.Forms.GroupBox
    $DistListRemoveUserBox1.Location = '40,40'
    $DistListRemoveUserBox1.size = '700,125'
    $DistListRemoveUserBox1.text = '1. Select a UserName and Distribution list from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $DistListRemoveUsertextLabel1 = New-Object System.Windows.Forms.Label
    $DistListRemoveUsertextLabel1.Location = '20,40'
    $DistListRemoveUsertextLabel1.size = '150,40'
    $DistListRemoveUsertextLabel1.Text = 'UserName:' 
    $DistListRemoveUsertextLabel2 = New-Object System.Windows.Forms.Label
    $DistListRemoveUsertextLabel2.Location = '20,80'
    $DistListRemoveUsertextLabel2.size = '150,40'
    $DistListRemoveUsertextLabel2.Text = 'Distribution list:' 
    ### Create group 1 box combo boxes. ###
    $DistListRemoveUserUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $DistListRemoveUserUserNameComboBox1.Location = '325,35'
    $DistListRemoveUserUserNameComboBox1.Size = '350, 310'
    $DistListRemoveUserUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $DistListRemoveUserUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $DistListRemoveUserUserNameComboBox1.Sorted = $true;
    $DistListRemoveUserUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $DistListRemoveUserUserNameComboBox1.DataSource = $UsernameList
    $DistListRemoveUserUserNameComboBox1.add_SelectedIndexChanged( { $DistListRemoveUserSelectedUserNametextLabel4.Text = "$($DistListRemoveUserUserNameComboBox1.SelectedItem.ToString())" })
    $DistListRemoveUserMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $DistListRemoveUserMBNameComboBox2.Location = '325,75'
    $DistListRemoveUserMBNameComboBox2.Size = '350, 350'
    $DistListRemoveUserMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $DistListRemoveUserMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $DistListRemoveUserMBNameComboBox2.Sorted = $true;
    $DistListRemoveUserMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $DistListRemoveUserMBNameComboBox2.DataSource = $DistributionLists 
    $DistListRemoveUserMBNameComboBox2.add_SelectedIndexChanged( { $DistListRemoveUserSelectedMailBoxNametextLabel6.Text = "$($DistListRemoveUserMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $DistListRemoveUserBox2 = New-Object System.Windows.Forms.GroupBox
    $DistListRemoveUserBox2.Location = '40,190'
    $DistListRemoveUserBox2.size = '700,125'
    $DistListRemoveUserBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $DistListRemoveUsertextLabel3 = New-Object System.Windows.Forms.Label
    $DistListRemoveUsertextLabel3.Location = '40,40'
    $DistListRemoveUsertextLabel3.size = '100,40'
    $DistListRemoveUsertextLabel3.Text = 'The User:' 
    $DistListRemoveUserSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $DistListRemoveUserSelectedUserNametextLabel4.Location = '30,80'
    $DistListRemoveUserSelectedUserNametextLabel4.Size = '200,40'
    $DistListRemoveUserSelectedUserNametextLabel4.ForeColor = 'Blue'
    $DistListRemoveUsertextLabel5 = New-Object System.Windows.Forms.Label
    $DistListRemoveUsertextLabel5.Location = '275,40'
    $DistListRemoveUsertextLabel5.size = '400,40'
    $DistListRemoveUsertextLabel5.Text = 'Will be removed from the Distribution list:'
    $DistListRemoveUserSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $DistListRemoveUserSelectedMailBoxNametextLabel6.Location = '350,80'
    $DistListRemoveUserSelectedMailBoxNametextLabel6.Size = '200,40'
    $DistListRemoveUserSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $DistListRemoveUserBox3 = New-Object System.Windows.Forms.GroupBox
    $DistListRemoveUserBox3.Location = '40,340'
    $DistListRemoveUserBox3.size = '700,30'
    $DistListRemoveUserBox3.text = '3. Click Ok to add User to list or Cancel:'
    $DistListRemoveUserBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $DistListRemoveUserForm.Close()
            $DistListRemoveUserForm.Dispose()
            Return DistributionListManagementForm })
    ### Add all the Form controls ### 
    $DistListRemoveUserForm.Controls.AddRange(@($DistListRemoveUserBox1, $DistListRemoveUserBox2, $DistListRemoveUserBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $DistListRemoveUserBox1.Controls.AddRange(@($DistListRemoveUsertextLabel1, $DistListRemoveUsertextLabel2, $DistListRemoveUserUserNameComboBox1, $DistListRemoveUserMBNameComboBox2))
    $DistListRemoveUserBox2.Controls.AddRange(@($DistListRemoveUsertextLabel3, $DistListRemoveUserSelectedUserNametextLabel4, $DistListRemoveUsertextLabel5, $DistListRemoveUserSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $DistListRemoveUserForm.AcceptButton = $OKButton
    $DistListRemoveUserForm.CancelButton = $CancelButton
    #### Activate the form ###
    $DistListRemoveUserForm.Add_Shown( { $DistListRemoveUserForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $DistListRemoveUserForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($DistListRemoveUserSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Distribution List - Remove User from a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListRemoveUserForm.Close()
            $DistListRemoveUserForm.Dispose()
            break
        }
        Elseif ($DistListRemoveUserSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Distribution List !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Distribution List - Remove User from a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListRemoveUserForm.Close()
            $DistListRemoveUserForm.Dispose()
            break
        }
        #####################################################################
        #########                get SamAccountNames            #############
        #####################################################################
        $UserSamAccountName = get-mailbox $($DistListRemoveUserUserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        $MailboxSamAccountname = Get-ADGroup -SearchBase 'OU=Distribution Lists,OU=Groups,OU=COURTS,DC=scotcourts,DC=local' -Filter "Name -eq '$($DistListRemoveUserSelectedMailBoxNametextLabel6.Text)'" | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName 
        #####################################################################
        #########        Check if User already is a member      #############
        #####################################################################
        $Status = Get-ADGroupMember -Identity ($MailboxSamAccountname) | Select-Object -ExpandProperty Name
        $Status1 = ($Status -match "$($DistListRemoveUserSelectedUserNametextLabel4.Text)").Count
        If ($Status1 -eq 0) {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user $($DistListRemoveUserUserNameComboBox1.SelectedItem.ToString()) isnt a member of the $($DistListRemoveUserMBNameComboBox2.SelectedItem.ToString()) distribution list.", 'Mailbox - Distribution List - Distribution List - Remove User from a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListRemoveUserForm.Close()
            $DistListRemoveUserForm.Dispose()
            Return DistributionListManagementForm
        }
        Else {
            #####################################################################
            ##   CHECK - continue if only 1 username in pipe if not exit       ##
            #####################################################################
            if (($UserSamAccountName | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            #
            #####################################################################
            ######                Remove user from list           ###############
            #####################################################################
            Remove-ADGroupMember -Identity $MailboxSamAccountname -Member $UserSamAccountName -Confirm:$false
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user $($DistListRemoveUserUserNameComboBox1.SelectedItem.ToString()) has been removed from Distribution list $($DistListRemoveUserMBNameComboBox2.SelectedItem.ToString())`n`nThe user needs to close and re-open Outlook.", 'Mailbox - Distribution List - Remove User from a List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: The User $($DistListRemoveUserUserNameComboBox1.SelectedItem.ToString()) has been removed from Distribution list $($DistListRemoveUserMBNameComboBox2.SelectedItem.ToString())." -Body 'The user needs to close and re-open Outlook.' -SmtpServer mail.scotcourts.local
            $DistListRemoveUserForm.Close()
            $DistListRemoveUserForm.Dispose()
            Return DistributionListManagementForm
        }
    }
}
#########################################################################
###  Completed -  Create SubForm 'Dist List - Remove User from a List ###
#########################################################################
#
######################################################################
# Create SubForm 'Distribution List - List current members of a List #
######################################################################
Function DistListListUsersForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $DistListListForm = New-Object System.Windows.Forms.Form
    $DistListListForm.width = 745
    $DistListListForm.height = 475
    $DistListListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DistListListForm.Controlbox = $false
    $DistListListForm.Icon = $Icon
    $DistListListForm.FormBorderStyle = 'Fixed3D'
    $DistListListForm.Text = 'Mailbox - Distribution List - List current members of a List.'
    $DistListListForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $DistListListBox1 = New-Object System.Windows.Forms.GroupBox
    $DistListListBox1.Location = '40,20'
    $DistListListBox1.size = '650,125'
    $DistListListBox1.text = '1. Select a Distribution List from the dropdown list:'
    ### Create group 1 box text labels. ###
    $DistListListtextLabel2 = New-Object System.Windows.Forms.Label
    $DistListListtextLabel2.Location = '20,50'
    $DistListListtextLabel2.size = '200,40'
    $DistListListtextLabel2.Text = 'Distribution List:' 
    ### Create group 1 box combo boxes. ###
    $DistListListMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $DistListListMBNameComboBox2.Location = '275,45'
    $DistListListMBNameComboBox2.Size = '350, 350'
    $DistListListMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $DistListListMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $DistListListMBNameComboBox2.Sorted = $true;
    $DistListListMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $DistListListMBNameComboBox2.DataSource = $DistributionLists
    $DistListListMBNameComboBox2.Add_SelectedIndexChanged( { $DistListListSelectedNametextLabel6.Text = "$($DistListListMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $DistListListBox2 = New-Object System.Windows.Forms.GroupBox
    $DistListListBox2.Location = '40,170'
    $DistListListBox2.size = '650,125'
    $DistListListBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $DistListListtextLabel3 = New-Object System.Windows.Forms.Label
    $DistListListtextLabel3.Location = '40,40'
    $DistListListtextLabel3.size = '400,40'
    $DistListListtextLabel3.Text = 'List current members of Distribution List:' 
    $DistListListSelectedNametextLabel6 = New-Object System.Windows.Forms.Label
    $DistListListSelectedNametextLabel6.Location = '100,80'
    $DistListListSelectedNametextLabel6.Size = '400,40'
    $DistListListSelectedNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $DistListListBox3 = New-Object System.Windows.Forms.GroupBox
    $DistListListBox3.Location = '40,320'
    $DistListListBox3.size = '650,30'
    $DistListListBox3.text = '3. Click Ok to List current members or Cancel:'
    $DistListListBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $DistListListForm.Close()
            $DistListListForm.Dispose()
            Return DistributionListManagementForm })
    ### Add all the Form controls ### 
    $DistListListForm.Controls.AddRange(@($DistListListBox1, $DistListListBox2, $DistListListBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $DistListListBox1.Controls.AddRange(@($DistListListtextLabel2, $DistListListMBNameComboBox2))
    $DistListListBox2.Controls.AddRange(@($DistListListtextLabel3, $DistListListtextLabel5, $DistListListSelectedNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $DistListListForm.AcceptButton = $OKButton
    $DistListListForm.CancelButton = $CancelButton
    #### Activate the form ###
    $DistListListForm.Add_Shown( { $DistListListForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $DistListListForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($DistListListSelectedNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Distribution list !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Distribution List - List current members of a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListListForm.Close()
            $DistListListForm.Dispose()
            break
        }
        #####################################################################
        ########               Get List of Names             ################ 
        #####################################################################
        $DistListName = $DistListListSelectedNametextLabel6.Text 
        Get-DistributionGroupMember -Identity $DistListName | Select-Object Name | Out-GridView -Title "List of current members of the $DistListName distribution list" -Wait
        $DistListListForm.Close()
        $DistListListForm.Dispose()
        Return DistributionListManagementForm
    }
}
######################################################################
### Completed -  Create SubForm 'Dist List - List current members  ###
######################################################################
#
######################################################################
####     Create SubForm  New Distribution List Sub Form           ####
######################################################################
Function DistListAddNew {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $DistListAddNewForm = New-Object System.Windows.Forms.Form
    $DistListAddNewForm.width = 780
    $DistListAddNewForm.height = 500
    $DistListAddNewForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DistListAddNewForm.Controlbox = $false
    $DistListAddNewForm.Icon = $Icon
    $DistListAddNewForm.FormBorderStyle = 'Fixed3D'
    $DistListAddNewForm.Text = 'Add New Distribution List'
    $DistListAddNewForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $DistListAddNewBox1 = New-Object System.Windows.Forms.GroupBox
    $DistListAddNewBox1.Location = '40,40'
    $DistListAddNewBox1.size = '700,125'
    $DistListAddNewBox1.text = '1. Enter the Distribution List Display Name and Email address:'
    ### Create group 1 box text labels. ###
    $DistListAddNewtextLabel1 = New-Object System.Windows.Forms.Label
    $DistListAddNewtextLabel1.Location = '20,40'
    $DistListAddNewtextLabel1.size = '300,40'
    $DistListAddNewtextLabel1.Text = 'Display Name:      (e.g. All Users Aberdeen)' 
    $DistListAddNewtextLabel2 = New-Object System.Windows.Forms.Label
    $DistListAddNewtextLabel2.Location = '20,79'
    $DistListAddNewtextLabel2.size = '350,40'
    $DistListAddNewtextLabel2.Text = 'Email address:     (e.g. allusersaberdeen with no spaces)' 
    ### Create group 1 box text boxes. ###
    $DistListAddNewUserNameComboBox1 = New-Object System.Windows.Forms.TextBox
    $DistListAddNewUserNameComboBox1.Location = '425,35'
    $DistListAddNewUserNameComboBox1.Size = '250,40'
    $DistListAddNewUserNameComboBox1.add_TextChanged( { $DistListAddNewSelectedUserNametextLabel4.Text = "$($DistListAddNewUserNameComboBox1.text)" })
    $DistListAddNewUserNameComboBox1.Add_TextChanged( { If ($This.Text -and $DistListAddNewMBNameComboBox2.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } })  
    $DistListAddNewMBNameComboBox2 = New-Object System.Windows.Forms.TextBox
    $DistListAddNewMBNameComboBox2.Location = '425,75'
    $DistListAddNewMBNameComboBox2.Size = '250,40'
    $DistListAddNewMBNameComboBox2.add_textChanged( { $DistListAddNewSelectedMailBoxNametextLabel6.Text = "$($DistListAddNewMBNameComboBox2.text)@scotcourts.gov.uk" })
    $DistListAddNewMBNameComboBox2.Add_TextChanged( { If ($This.Text -and $DistListAddNewUserNameComboBox1.Text) { $OKButton.Enabled = $True }Else { $OKButton.Enabled = $False } }) 
    ### Create group 2 box in form. ###
    $DistListAddNewBox2 = New-Object System.Windows.Forms.GroupBox
    $DistListAddNewBox2.Location = '40,190'
    $DistListAddNewBox2.size = '700,125'
    $DistListAddNewBox2.text = '2. Check the details below are correct before proceeding:'
    ### Create group 2 box text labels.
    $DistListAddNewtextLabel3 = New-Object System.Windows.Forms.Label
    $DistListAddNewtextLabel3.Location = '20,40'
    $DistListAddNewtextLabel3.size = '350,40'
    $DistListAddNewtextLabel3.Text = 'Distribution List will appear in Global Adress List as:' 
    $DistListAddNewSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $DistListAddNewSelectedUserNametextLabel4.Location = '40,80'
    $DistListAddNewSelectedUserNametextLabel4.Size = '250,40'
    $DistListAddNewSelectedUserNametextLabel4.ForeColor = 'Blue'
    $DistListAddNewtextLabel5 = New-Object System.Windows.Forms.Label
    $DistListAddNewtextLabel5.Location = '430,40'
    $DistListAddNewtextLabel5.size = '200,40'
    $DistListAddNewtextLabel5.Text = 'With the email address:'
    $DistListAddNewSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $DistListAddNewSelectedMailBoxNametextLabel6.Location = '380,80'
    $DistListAddNewSelectedMailBoxNametextLabel6.Size = '400,40'
    $DistListAddNewSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $DistListAddNewBox3 = New-Object System.Windows.Forms.GroupBox
    $DistListAddNewBox3.Location = '40,340'
    $DistListAddNewBox3.size = '700,30'
    $DistListAddNewBox3.text = '3. Click Ok to add New Distribution List or Cancel:'
    $DistListAddNewBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $DistListAddNewForm.Close()
            $DistListAddNewForm.Dispose()
            Return DistributionListManagementForm })
    ### Add all the Form controls ### 
    $DistListAddNewForm.Controls.AddRange(@($DistListAddNewBox1, $DistListAddNewBox2, $DistListAddNewBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $DistListAddNewBox1.Controls.AddRange(@($DistListAddNewtextLabel1, $DistListAddNewtextLabel2, $DistListAddNewUserNameComboBox1, $DistListAddNewMBNameComboBox2))
    $DistListAddNewBox2.Controls.AddRange(@($DistListAddNewtextLabel3, $DistListAddNewSelectedUserNametextLabel4, $DistListAddNewtextLabel5, $DistListAddNewSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $DistListAddNewForm.AcceptButton = $OKButton
    $DistListAddNewForm.CancelButton = $CancelButton
    #### Activate the form ###
    $DistListAddNewForm.Add_Shown( { $DistListAddNewForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $DistListAddNewForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or mailbox     ################ 
        #####################################################################
        if ($DistListAddNewSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to type a Display Name !!!!!  Trying to enter blank fields is never a good idea.", 'Add New Distribution List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListAddNewForm.Close()
            $DistListAddNewForm.Dispose()
            break
        }
        Elseif ($DistListAddNewSelectedMailBoxNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to type an Email address name !!!!!  Trying to enter blank fields is never a good idea.", 'Add New Distribution List.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DistListAddNewForm.Close()
            $DistListAddNewForm.Dispose()
            break
        }
        #####################################################################
        #########   Check if email address is already in use    #############
        #####################################################################
        $DisplayName = $DistListAddNewUserNameComboBox1.Text
        $EmailName = $DistListAddNewMBNameComboBox2.Text   
        $ListEmailAddress = Get-ADObject -Filter "mail -eq '$EmailName@scotcourts.gov.uk'" | Measure-Object count 
        If ($Null -ne $ListEmailAddress) {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The Distribution list - $DisplayName - can not be added because the email address $EmailName@scotcourts.gov.uk is currently in use on another list  Please use a name/email address thats not currently in use.", 'ERROR - CANT ADD NEW DISTRIBUTION LIST', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return DistListAddNew
        }
        Else {
            #####################################################################
            #    CHECK - continue if only 1 EmailName in pipe if not exit       #
            #####################################################################
            if (($EmailName | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            # 
            #####################################################################
            ######           Add New Distribution List            ###############
            #####################################################################
            New-DistributionGroup -Name $DisplayName -OrganizationalUnit 'ou=Distribution Lists,ou=Groups,ou=SCTS,dc=scotcourts,dc=local' -SamAccountName $EmailName -Type Distribution
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("New Distribution List - $DisplayName - has been added  The members now need to be added to the list.", 'Add New Distribution List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            #############   Reload Distribution lists       ####################
            #####################################################################
            #####################################################################
            ###  Create form to pause for 7 sec  ###
            Add-Type -AssemblyName System.Windows.Forms
            ### Build Form ###
            $objForm = New-Object System.Windows.Forms.Form
            $objForm.Text = 'Add New User Account'
            $objForm.Size = New-Object System.Drawing.Size(450, 270)
            $objForm.StartPosition = 'CenterScreen'
            $objForm.Controlbox = $false
            #### Add Label ###
            $objLabel = New-Object System.Windows.Forms.Label
            $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
            $objLabel.Size = New-Object System.Drawing.Size(300, 120)
            $objLabel.Text = 'The Distribution Lists are being reloaded from AD with the  ew List you have just added.     Please Wait. ..............'
            $objForm.Controls.Add($objLabel)
            ### Show the form ###
            $objForm.Show() | Out-Null
            ### wait 7 seconds ###
            Start-Sleep -Seconds 7
            ### destroy form ###
            $objForm.Close() | Out-Null
            ### update dist lists with new added list ###
            function changeit {  
                $script:Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name  
            }  
            changeit 
            #####################################################################
            #############   Send email to helpdesk    ###########################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New Distribution List - $DisplayName - has been added" -Body "New Distribution List - $DisplayName - has been added  The members now need to be added to the list." -SmtpServer mail.scotcourts.local
            $DistListAddNewForm.Close()
            $DistListAddNewForm.Dispose()
            Return DistributionListManagementForm
        }
    }
}
#########################################################################
#####       Completed -  Create New Distribution List form          #####
######################################################################### 
#
#########################################################################################################
##########   Completed - Create 'Shared Mailbox - Distribution List Management' Sub Forms   #############
#########################################################################################################
#
#########################################################################################################
###########             Create the 'Shared Mailbox - Management' form      ############################## 
#########################################################################################################
Function SharedMailboxManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details of the form.
    $SharedMailboxForm = New-Object System.Windows.Forms.Form
    $SharedMailboxForm.width = 750
    $SharedMailboxForm.height = 550
    $SharedMailboxForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $SharedMailboxForm.MinimizeBox = $False
    $SharedMailboxForm.MaximizeBox = $False
    $SharedMailboxForm.FormBorderStyle = 'Fixed3D'
    $SharedMailboxForm.Text = 'Shared Mailbox - Management.'
    $SharedMailboxForm.Icon = $Icon
    $SharedMailboxForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain the radio buttons.
    $SharedMailBoxBox = New-Object System.Windows.Forms.GroupBox
    $SharedMailBoxBox.Location = '40,30'
    $SharedMailBoxBox.size = '650,410'
    $SharedMailBoxBox.text = 'Select an option.'
    # Create the radio buttons
    $SharedMailBoxButton1 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton1.Location = '20,20'
    $SharedMailBoxButton1.size = '600,40'
    $SharedMailBoxButton1.Checked = $true 
    $SharedMailBoxButton1.Text = 'Add - Full Access permissions for a User.'
    $SharedMailBoxButton2 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton2.Location = '20,50'
    $SharedMailBoxButton2.size = '600,40'
    $SharedMailBoxButton2.Checked = $false
    $SharedMailBoxButton2.Text = 'Add - Send On Behalf Of permissions for a User.'
    $SharedMailBoxButton3 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton3.Location = '20,90'
    $SharedMailBoxButton3.size = '600,40'
    $SharedMailBoxButton3.Checked = $false
    $SharedMailBoxButton3.Text = 'Check - Mailbox current Full Access permissions.'
    $SharedMailBoxButton4 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton4.Location = '20,120'
    $SharedMailBoxButton4.size = '600,40'
    $SharedMailBoxButton4.Checked = $false
    $SharedMailBoxButton4.Text = 'Check - Mailbox current Send On Behalf Of permissions.'
    $SharedMailBoxButton5 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton5.Location = '20,160'
    $SharedMailBoxButton5.size = '600,40'
    $SharedMailBoxButton5.Checked = $false
    $SharedMailBoxButton5.Text = 'Remove - Full Access permissions for a User.'
    $SharedMailBoxButton6 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton6.Location = '20,190'
    $SharedMailBoxButton6.size = '600,40'
    $SharedMailBoxButton6.Checked = $false
    $SharedMailBoxButton6.Text = 'Remove - Send On Behalf Of permissions for a User.'
    $SharedMailBoxButton7 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton7.Location = '20,240'
    $SharedMailBoxButton7.size = '600,40'
    $SharedMailBoxButton7.Checked = $false
    $SharedMailBoxButton7.Text = 'Shared Mailbox Log On - Access for Autoreply Rules, Out of Office, Check emails etc.'
    $SharedMailBoxButton8 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton8.Location = '20,290'
    $SharedMailBoxButton8.size = '600,40'
    $SharedMailBoxButton8.Checked = $false
    $SharedMailBoxButton8.Text = 'New Courts Shared mailbox - Add new Courts Shared mailbox.'
    $SharedMailBoxButton9 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton9.Location = '20,320'
    $SharedMailBoxButton9.size = '600,40'
    $SharedMailBoxButton9.Checked = $false
    $SharedMailBoxButton9.Text = 'New Tribunals Shared mailbox - Add new Tribunals Shared mailbox.'
    $SharedMailBoxButton10 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxButton10.Location = '20,360'
    $SharedMailBoxButton10.size = '600,40'
    $SharedMailBoxButton10.Checked = $false
    $SharedMailBoxButton10.Text = 'Add copy of sent email to shared mailbox sent folder.'    # Add an OK button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,440'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,440'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $SharedMailboxForm.Close()
            $SharedMailboxForm.Dispose()
            Return MainForm })
    # Add all the Form controls on one line 
    $SharedMailboxForm.Controls.AddRange(@($SharedMailBoxBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $SharedMailBoxBox.Controls.AddRange(@($SharedMailBoxButton1, $SharedMailBoxButton2, $SharedMailBoxButton3, $SharedMailBoxButton4, $SharedMailBoxButton5, $SharedMailBoxButton6, $SharedMailBoxButton7, $SharedMailBoxButton8, $SharedMailBoxButton9, $SharedMailBoxButton10))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $SharedMailboxForm.AcceptButton = $OKButton
    $SharedMailboxForm.CancelButton = $CancelButton
    # Activate the form
    $SharedMailboxForm.Add_Shown( { $SharedMailboxForm.Activate() })    
    # Get the results from the button click
    $dialogResult = $SharedMailboxForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        # Check the current state of each radio button and respond
        if ($SharedMailBoxButton1.Checked) { AddMBFullAccessPermissionForm }
        elseif ($SharedMailBoxButton2.Checked) { AddMBSendBehalfOfAccessPermissionForm }
        elseif ($SharedMailBoxButton3.Checked) { CheckMBFullAccessPermissionForm }
        elseif ($SharedMailBoxButton4.Checked) { CheckMBSendOnBehalfPermissionForm }
        elseif ($SharedMailBoxButton5.Checked) { RemoveMBFullAccessPermissionForm }
        elseif ($SharedMailBoxButton6.Checked) { RemoveMBSendOnBehalfPermissionForm }
        elseif ($SharedMailBoxButton7.Checked) { SharedLogOnForm }
        elseif ($SharedMailBoxButton8.Checked) { AddNewSharedMBForm }
        elseif ($SharedMailBoxButton9.Checked) { AddNewTribunalSharedMBForm }
        elseif ($SharedMailBoxButton10.Checked) { CopySharedMBForm }
    }
}
#########################################################################################################
###########        Completed - Create the 'Shared Mailbox - Management' form            ################# 
#########################################################################################################
#
#########################################################################################################
###########        Create the 'User Mailbox - Out Of Office management' form      ####################### 
#########################################################################################################
Function OutOfficeManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details of the form.
    $OutOfficeManagementForm = New-Object System.Windows.Forms.Form
    $OutOfficeManagementForm.width = 750
    $OutOfficeManagementForm.height = 450
    $OutOfficeManagementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $OutOfficeManagementForm.MinimizeBox = $False
    $OutOfficeManagementForm.MaximizeBox = $False
    $OutOfficeManagementForm.FormBorderStyle = 'Fixed3D'
    $OutOfficeManagementForm.Text = 'User Mailbox - Out Of Office management.'
    $OutOfficeManagementForm.Icon = $Icon
    $OutOfficeManagementForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain the radio buttons.
    $OutOfficeBox = New-Object System.Windows.Forms.GroupBox
    $OutOfficeBox.Location = '40,30'
    $OutOfficeBox.size = '650,300'
    $OutOfficeBox.text = 'Select an option.'
    # Create the radio buttons.
    $OutOfficeButton1 = New-Object System.Windows.Forms.RadioButton
    $OutOfficeButton1.Location = '20,40'
    $OutOfficeButton1.size = '600,40'
    $OutOfficeButton1.Checked = $true 
    $OutOfficeButton1.Text = 'Out Of Office - Check current status:'
    $OutOfficeButton2 = New-Object System.Windows.Forms.RadioButton
    $OutOfficeButton2.Location = '20,80'
    $OutOfficeButton2.size = '600,40'
    $OutOfficeButton2.Checked = $false 
    $OutOfficeButton2.Text = 'Out Of Office - Turn On:'
    $OutOfficeButton3 = New-Object System.Windows.Forms.RadioButton
    $OutOfficeButton3.Location = '20,120'
    $OutOfficeButton3.size = '600,40'
    $OutOfficeButton3.Checked = $false
    $OutOfficeButton3.Text = 'Out Of Office - Turn Off:'
    $OutOfficeButton4 = New-Object System.Windows.Forms.RadioButton
    $OutOfficeButton4.Location = '20,160'
    $OutOfficeButton4.size = '600,40'
    $OutOfficeButton4.Checked = $false
    $OutOfficeButton4.Text = 'Out Of Office - Add message and turn Out of Office On:'
    $OutOfficeButton5 = New-Object System.Windows.Forms.RadioButton
    $OutOfficeButton5.Location = '20,200'
    $OutOfficeButton5.size = '600,40'
    $OutOfficeButton5.Checked = $false
    $OutOfficeButton5.Text = 'Out Of Office - Check current message:'
    # Add an OK button.
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,350'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,350'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $OutOfficeManagementForm.Close()
            $OutOfficeManagementForm.Dispose()
            MainForm })
    # Add all the Form controls on one line 
    $OutOfficeManagementForm.Controls.AddRange(@($OutOfficeBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $OutOfficeBox.Controls.AddRange(@($OutOfficeButton1, $OutOfficeButton2, $OutOfficeButton3, $OutOfficeButton4, $OutOfficeButton5))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $OutOfficeManagementForm.AcceptButton = $OKButton
    $OutOfficeManagementForm.CancelButton = $CancelButton
    # Activate the form
    $OutOfficeManagementForm.Add_Shown( { $OutOfficeManagementForm.Activate() })    
    # Get the results from the button click
    $Result = $OutOfficeManagementForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        # Check the current state of each radio button and respond.
        if ($OutOfficeButton1.Checked) { OutOfficeCheckForm }
        elseif ($OutOfficeButton2.Checked) { OutOfficeTurnOnForm }
        elseif ($OutOfficeButton3.Checked) { OutOfficeTurnOffForm }
        elseif ($OutOfficeButton4.Checked) { OutOfficeAddTurnOnForm }
        elseif ($OutOfficeButton5.Checked) { OutOfficeCheckMessageForm }
    }
}
#########################################################################################################
#########      Completed -   Create the 'User Mailbox - Out Of Office management form    ################ 
#########################################################################################################
#
#########################################################################################################
#################        Create the 'Shared Calendar Management' form     ############################### 
#########################################################################################################
Function SharedCalendarManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details of the form.
    $SharedCalendarManagementForm = New-Object System.Windows.Forms.Form
    $SharedCalendarManagementForm.width = 750
    $SharedCalendarManagementForm.height = 450
    $SharedCalendarManagementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $SharedCalendarManagementForm.MinimizeBox = $False
    $SharedCalendarManagementForm.MaximizeBox = $False
    $SharedCalendarManagementForm.FormBorderStyle = 'Fixed3D'
    $SharedCalendarManagementForm.Text = 'Shared Calendar - Management.'
    $SharedCalendarManagementForm.Icon = $Icon
    $SharedCalendarManagementForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain the radio buttons.
    $SharedMailBoxCalendarBox = New-Object System.Windows.Forms.GroupBox
    $SharedMailBoxCalendarBox.Location = '40,30'
    $SharedMailBoxCalendarBox.size = '650,300'
    $SharedMailBoxCalendarBox.text = 'Select an option.'
    # Create the radio buttons.
    $SharedMailBoxCalendarButton1 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxCalendarButton1.Location = '20,40'
    $SharedMailBoxCalendarButton1.size = '600,40'
    $SharedMailBoxCalendarButton1.Checked = $true 
    $SharedMailBoxCalendarButton1.Text = 'Add - Owner permissions for a User (full access).'
    $SharedMailBoxCalendarButton2 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxCalendarButton2.Location = '20,80'
    $SharedMailBoxCalendarButton2.size = '600,40'
    $SharedMailBoxCalendarButton2.Checked = $false 
    $SharedMailBoxCalendarButton2.Text = 'Add - Reviewer permissions for a User (read only).'
    $SharedMailBoxCalendarButton3 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxCalendarButton3.Location = '20,120'
    $SharedMailBoxCalendarButton3.size = '600,40'
    $SharedMailBoxCalendarButton3.Checked = $false
    $SharedMailBoxCalendarButton3.Text = 'Check - Calendar current permissions.'
    $SharedMailBoxCalendarButton4 = New-Object System.Windows.Forms.RadioButton
    $SharedMailBoxCalendarButton4.Location = '20,160'
    $SharedMailBoxCalendarButton4.size = '600,40'
    $SharedMailBoxCalendarButton4.Checked = $false
    $SharedMailBoxCalendarButton4.Text = 'Remove - Remove Calendar permissions for a User.'
    # Add an OK button.
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,350'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,350'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $SharedCalendarManagementForm.Close()
            $SharedCalendarManagementForm.Dispose()
            Return MainForm })
    # Add all the Form controls on one line 
    $SharedCalendarManagementForm.Controls.AddRange(@($SharedMailBoxCalendarBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $SharedMailBoxCalendarBox.Controls.AddRange(@($SharedMailBoxCalendarButton1, $SharedMailBoxCalendarButton2, $SharedMailBoxCalendarButton3, $SharedMailBoxCalendarButton4))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $SharedCalendarManagementForm.AcceptButton = $OKButton
    $SharedCalendarManagementForm.CancelButton = $CancelButton
    # Activate the form
    $SharedCalendarManagementForm.Add_Shown( { $SharedCalendarManagementForm.Activate() })    
    # Get the results from the button click
    $Result = $SharedCalendarManagementForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        # Check the current state of each radio button and respond.
        if ($SharedMailBoxCalendarButton1.Checked) { AddCalOwnerPermissionForm }
        elseif ($SharedMailBoxCalendarButton2.Checked) { AddCalReviewerPermissionForm }
        elseif ($SharedMailBoxCalendarButton3.Checked) { CheckCalPermAccessPermissionForm }
        elseif ($SharedMailBoxCalendarButton4.Checked = $true) { RemoveCalendarPermAccessPermissionForm }
    }
}
#########################################################################################################
#################       Completed - Create the 'Shared Calendar Management form'     #################### 
#########################################################################################################
#
#########################################################################################################
###########        Create the 'DistributionList Management' form     #################################### 
#########################################################################################################
Function DistributionListManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details of the form.
    $DistributionListManagementForm = New-Object System.Windows.Forms.Form
    $DistributionListManagementForm.width = 750
    $DistributionListManagementForm.height = 450
    $DistributionListManagementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DistributionListManagementForm.MinimizeBox = $False
    $DistributionListManagementForm.MaximizeBox = $False
    $DistributionListManagementForm.FormBorderStyle = 'Fixed3D'
    $DistributionListManagementForm.Text = 'Distribution List - Management.'
    $DistributionListManagementForm.Icon = $Icon
    $DistributionListManagementForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain the radio buttons.
    $DistributionListBox = New-Object System.Windows.Forms.GroupBox
    $DistributionListBox.Location = '40,30'
    $DistributionListBox.size = '650,300'
    $DistributionListBox.text = 'Select an option.'
    # Create the radio buttons.
    $DistributionListButton1 = New-Object System.Windows.Forms.RadioButton
    $DistributionListButton1.Location = '20,40'
    $DistributionListButton1.size = '600,40'
    $DistributionListButton1.Checked = $true 
    $DistributionListButton1.Text = 'Distribution List - Add User to a List:'
    $DistributionListButton2 = New-Object System.Windows.Forms.RadioButton
    $DistributionListButton2.Location = '20,80'
    $DistributionListButton2.size = '600,40'
    $DistributionListButton2.Checked = $false 
    $DistributionListButton2.Text = 'Distribution List - Remove User from a List:'
    $DistributionListButton3 = New-Object System.Windows.Forms.RadioButton
    $DistributionListButton3.Location = '20,120'
    $DistributionListButton3.size = '600,40'
    $DistributionListButton3.Checked = $false
    $DistributionListButton3.Text = 'Distribution List - List current members of a List:'
    $DistributionListButton4 = New-Object System.Windows.Forms.RadioButton
    $DistributionListButton4.Location = '20,200'
    $DistributionListButton4.size = '600,40'
    $DistributionListButton4.Checked = $false
    $DistributionListButton4.Text = 'Distribution List - Add New Distribution List:'
    # Add an OK button.
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,350'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,350'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $DistributionListManagementForm.Close()
            $DistributionListManagementForm.Dispose()
            Return MainForm })
    # Add all the Form controls on one line 
    $DistributionListManagementForm.Controls.AddRange(@($DistributionListBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $DistributionListBox.Controls.AddRange(@($DistributionListButton1, $DistributionListButton2, $DistributionListButton3, $DistributionListButton4))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $DistributionListManagementForm.AcceptButton = $OKButton
    $DistributionListManagementForm.CancelButton = $CancelButton
    # Activate the form
    $DistributionListManagementForm.Add_Shown( { $DistributionListManagementForm.Activate() })    
    # Get the results from the button click
    $Result = $DistributionListManagementForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        # Check the current state of each radio button and respond.
        if ($DistributionListButton1.Checked) { DistributionListAddUserForm }
        elseif ($DistributionListButton2.Checked) { DistributionListRemoveUsererForm }
        elseif ($DistributionListButton3.Checked) { DistListListUsersForm }
        elseif ($DistributionListButton4.Checked) { DistListAddNew }
    }
}
#########################################################################################################
###########      Completed -   Create the 'DistributionList Management' form     ######################## 
#########################################################################################################
#
#########################################################################################################
###########          Create the Disabled User Account form          ##################################### 
#########################################################################################################
Function DisabledUserManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $DisabledUserManagementForm = New-Object System.Windows.Forms.Form
    $DisabledUserManagementForm.width = 780
    $DisabledUserManagementForm.height = 500
    $DisabledUserManagementForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $DisabledUserManagementForm.Controlbox = $false
    $DisabledUserManagementForm.Icon = $Icon
    $DisabledUserManagementForm.FormBorderStyle = 'Fixed3D'
    $DisabledUserManagementForm.Text = 'Mailbox - Set Disabled Users mailbox only accept from helpdesk & Hide from Address List.'
    $DisabledUserManagementForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $HideAcceptBox1 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox1.Location = '40,40'
    $HideAcceptBox1.size = '700,125'
    $HideAcceptBox1.text = '1. Select a UserName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $HideAccepttextLabel1 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel1.Location = '20,40'
    $HideAccepttextLabel1.size = '150,40'
    $HideAccepttextLabel1.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $HideAcceptUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $HideAcceptUserNameComboBox1.Location = '325,35'
    $HideAcceptUserNameComboBox1.Size = '350, 310'
    $HideAcceptUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $HideAcceptUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $HideAcceptUserNameComboBox1.Sorted = $true;
    $HideAcceptUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $HideAcceptUserNameComboBox1.DataSource = $UsernameList
    $HideAcceptUserNameComboBox1.add_SelectedIndexChanged( { $HideAcceptSelectedUserNametextLabel4.Text = "$($HideAcceptUserNameComboBox1.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $HideAcceptBox2 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox2.Location = '40,190'
    $HideAcceptBox2.size = '700,125'
    $HideAcceptBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $HideAccepttextLabel3 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel3.Location = '40,40'
    $HideAccepttextLabel3.size = '100,40'
    $HideAccepttextLabel3.Text = 'The User:' 
    $HideAcceptSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $HideAcceptSelectedUserNametextLabel4.Location = '30,80'
    $HideAcceptSelectedUserNametextLabel4.Size = '200,40'
    $HideAcceptSelectedUserNametextLabel4.ForeColor = 'Blue'
    $HideAccepttextLabel5 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel5.Location = '275,40'
    $HideAccepttextLabel5.size = '400,40'
    $HideAccepttextLabel5.Text = 'Will be hidden from the Address List and can only receive emails from the IT helpdesk:'
    $HideAcceptSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $HideAcceptSelectedMailBoxNametextLabel6.Location = '350,80'
    $HideAcceptSelectedMailBoxNametextLabel6.Size = '200,40'
    $HideAcceptSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $HideAcceptBox3 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox3.Location = '40,340'
    $HideAcceptBox3.size = '700,30'
    $HideAcceptBox3.text = '3. Click Ok to confirm or Cancel:'
    $HideAcceptBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $DisabledUserManagementForm.Close()
            $DisabledUserManagementForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $DisabledUserManagementForm.Controls.AddRange(@($HideAcceptBox1, $HideAcceptBox2, $HideAcceptBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $HideAcceptBox1.Controls.AddRange(@($HideAccepttextLabel1, $HideAcceptUserNameComboBox1))
    $HideAcceptBox2.Controls.AddRange(@($HideAccepttextLabel3, $HideAcceptSelectedUserNametextLabel4, $HideAccepttextLabel5, $HideAcceptSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $DisabledUserManagementForm.AcceptButton = $OKButton
    $DisabledUserManagementForm.CancelButton = $CancelButton
    #### Activate the form ###
    $DisabledUserManagementForm.Add_Shown( { $DisabledUserManagementForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $DisabledUserManagementForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($HideAcceptSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mailbox - Set Disabled Users mailbox only accept from helpdesk & Hide from Address List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $DisabledUserManagementForm.Close()
            $DisabledUserManagementForm.Dispose()
            break
        }
        Else {
            #####################################################################
            ########         Check user has email address        ################ 
            #####################################################################
            $MailboxCheck = Get-ADUser -Filter { DisplayName -eq $HideAcceptSelectedUserNametextLabel4.Text } -Properties * | Select-Object EmailAddress
            If ($Null -eq $MailboxCheck.EmailAddress) {
                [System.Windows.Forms.MessageBox]::Show("The selected User does not have a mailbox  The options cannot be set.", 'Mailbox - Set Disabled Users mailbox only accept from helpdesk & Hide from Address List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $DisabledUserManagementForm.Close()
                $DisabledUserManagementForm.Dispose()
                Return DisabledUserManagementForm
            }
            #####################################################################
            ###  get mailbox primary smtpaddress from mailbox display name:  ####
            #####################################################################
            $MailBoxPrimarySMTPAddress = get-mailbox $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            #  Hide emailaddress from Address list & accept only from helpdesk  #
            #####################################################################
            Get-Mailbox $MailBoxPrimarySMTPAddress | Set-Mailbox -AcceptMessagesOnlyFrom 'helpdesk' -HiddenFromAddressListsEnabled $true
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            #Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) ) has been removed from the address list  and can now only accept emails from the helpdesk.", 'Mailbox - Set Disabled Users mailbox only accept from helpdesk & Hide from Address List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            ###############     send email to helpdesk    #######################
            #####################################################################
            $DateToDelete = Get-Date -Date $(Get-Date).adddays(30) -Format D
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "Delete user account $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) on $DateToDelete" -Body "The mailbox for $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) has been hidden from the Address Lists  It is now only accepting incoming email from the IT helpdesk  The user account should be deleted on $DateToDelete" -SmtpServer mail.scotcourts.local
            $DisabledUserManagementForm.Close()
            $DisabledUserManagementForm.Dispose()
            Return MainForm
        }
    }
}
#########################################################################################################
###########       Completed -    Create the Disabled User Account form    ############################### 
#########################################################################################################
#
#########################################################################################################
###########           Create the Mobile device reset form           ##################################### 
#########################################################################################################
Function MobileResetform {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $MobileResetform = New-Object System.Windows.Forms.Form
    $MobileResetform.width = 780
    $MobileResetform.height = 500
    $MobileResetform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $MobileResetform.Controlbox = $false
    $MobileResetform.Icon = $Icon
    $MobileResetform.FormBorderStyle = 'Fixed3D'
    $MobileResetform.Text = 'Mobile device reset form.'
    $MobileResetform.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $HideAcceptBox1 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox1.Location = '40,40'
    $HideAcceptBox1.size = '700,125'
    $HideAcceptBox1.text = '1. Select a UserName from the dropdown list:'
    ### Create group 1 box text labels. ###
    $HideAccepttextLabel1 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel1.Location = '20,40'
    $HideAccepttextLabel1.size = '150,40'
    $HideAccepttextLabel1.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $HideAcceptUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $HideAcceptUserNameComboBox1.Location = '325,35'
    $HideAcceptUserNameComboBox1.Size = '350, 310'
    $HideAcceptUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $HideAcceptUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $HideAcceptUserNameComboBox1.Sorted = $true;
    $HideAcceptUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $HideAcceptUserNameComboBox1.DataSource = $UsernameList
    $HideAcceptUserNameComboBox1.add_SelectedIndexChanged( { $HideAcceptSelectedUserNametextLabel4.Text = "$($HideAcceptUserNameComboBox1.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $HideAcceptBox2 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox2.Location = '40,190'
    $HideAcceptBox2.size = '700,125'
    $HideAcceptBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $HideAccepttextLabel3 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel3.Location = '40,40'
    $HideAccepttextLabel3.size = '100,40'
    $HideAccepttextLabel3.Text = 'The User:' 
    $HideAcceptSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $HideAcceptSelectedUserNametextLabel4.Location = '30,80'
    $HideAcceptSelectedUserNametextLabel4.Size = '200,40'
    $HideAcceptSelectedUserNametextLabel4.ForeColor = 'Blue'
    $HideAccepttextLabel5 = New-Object System.Windows.Forms.Label
    $HideAccepttextLabel5.Location = '275,40'
    $HideAccepttextLabel5.size = '400,40'
    $HideAccepttextLabel5.Text = 'Will be enabled to allow mobile device to access emails while removing any previous devices.'
    $HideAcceptSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
    $HideAcceptSelectedMailBoxNametextLabel6.Location = '350,80'
    $HideAcceptSelectedMailBoxNametextLabel6.Size = '200,40'
    $HideAcceptSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $HideAcceptBox3 = New-Object System.Windows.Forms.GroupBox
    $HideAcceptBox3.Location = '40,340'
    $HideAcceptBox3.size = '700,30'
    $HideAcceptBox3.text = '3. Click Ok to confirm or Cancel:'
    $HideAcceptBox3.button
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $MobileResetform.Close()
            $MobileResetform.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $MobileResetform.Controls.AddRange(@($HideAcceptBox1, $HideAcceptBox2, $HideAcceptBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $HideAcceptBox1.Controls.AddRange(@($HideAccepttextLabel1, $HideAcceptUserNameComboBox1))
    $HideAcceptBox2.Controls.AddRange(@($HideAccepttextLabel3, $HideAcceptSelectedUserNametextLabel4, $HideAccepttextLabel5, $HideAcceptSelectedMailBoxNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $MobileResetform.AcceptButton = $OKButton
    $MobileResetform.CancelButton = $CancelButton
    #### Activate the form ###
    $MobileResetform.Add_Shown( { $MobileResetform.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $MobileResetform.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########           Don't accept null mailbox         ################ 
        #####################################################################
        if ($HideAcceptSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Mobile device reset', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $MobileResetform.Close()
            $MobileResetform.Dispose()
            break
        }
        Else {
            #####################################################################
            ########         Check user has email address        ################ 
            #####################################################################
            $MailboxCheck = Get-ADUser -Filter { DisplayName -eq $HideAcceptSelectedUserNametextLabel4.Text } -Properties * | Select-Object EmailAddress
            If ($Null -eq $MailboxCheck.EmailAddress) {
                [System.Windows.Forms.MessageBox]::Show("The selected User does not have a mailbox  The options cannot be set.", 'Mobile device reset', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $MobileResetform.Close()
                $MobileResetform.Dispose()
                Return MobileResetform
            }
            #####################################################################
            ###  get mailbox primary smtpaddress from mailbox display name:  ####
            #####################################################################
            $MailBoxPrimarySMTPAddress = get-mailbox $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) | Select-Object primarysmtpaddress | Select-Object -ExpandProperty PrimarySMTPAddress  
            #####################################################################
            #  CHECK - continue if only 1 email address is in pipe if not exit  #
            #####################################################################
            if (($MailBoxPrimarySMTPAddress | Measure-Object).count -ne 1) { Mainform }
            #####################################################################
            ########################  Enable activeSync  ########################
            #####################################################################
            Get-Mailbox $MailBoxPrimarySMTPAddress | Set-CASMailbox -ActiveSyncEnabled $True
            #####################################################################
            ######## Remove any current mobile devices set on ActiveSync ########
            #####################################################################
            Get-ActiveSyncDevice -Mailbox $MailBoxPrimarySMTPAddress | Remove-ActiveSyncDevice
            #####################################################################
            ##################  Message complete message  #######################
            #####################################################################
            #Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) ) has been enabled for ActiveSync & any old device removed.", 'Mobile device reset', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            #####################################################################
            ###############    send email to helpdesk     #######################
            #####################################################################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -From $env:UserName@scotcourts.gov.uk -Subject "Mobile device reset on User account $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) on $DateToDelete" -Body "The mailbox for $($HideAcceptUserNameComboBox1.SelectedItem.ToString()) has been reset for mobile devices." -SmtpServer mail.scotcourts.local
            $MobileResetform.Close()
            $MobileResetform.Dispose()
            Return MainForm
        }
    }
}
#########################################################################################################
###########        Completed -    Create the Mobile device reset form     ############################### 
#########################################################################################################
#
#endregion Subforms
#region Mainform
#########################################################################################################
###########    Create the 'Exchange 2103 User Management main form' form   ############################## 
#########################################################################################################
Function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details size of your form
    $Ex2013ManForm = New-Object System.Windows.Forms.Form
    $Ex2013ManForm.width = 780
    $Ex2013ManForm.height = 500
    $Ex2013ManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Ex2013ManForm.MinimizeBox = $False
    $Ex2013ManForm.MaximizeBox = $False
    $Ex2013ManForm.FormBorderStyle = 'Fixed3D'
    $Ex2013ManForm.Text = 'Exchange 2013 Management Main Form v1.54'
    $Ex2013ManForm.Icon = $Icon
    $Ex2013ManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain your radio buttons
    $Ex2013ManGroupBox = New-Object System.Windows.Forms.GroupBox
    $Ex2013ManGroupBox.Location = '40,50'
    $Ex2013ManGroupBox.size = '700,320'
    $Ex2013ManGroupBox.text = 'Select an option.'
    # Create the collection of radio buttons
    $Ex2013ManGroupBoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton1.Location = '20,30'
    $Ex2013ManGroupBoxRadioButton1.size = '600,40'
    $Ex2013ManGroupBoxRadioButton1.Checked = $true 
    $Ex2013ManGroupBoxRadioButton1.Text = 'Shared Mailbox - management.'
    $Ex2013ManGroupBoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton2.Location = '20,70'
    $Ex2013ManGroupBoxRadioButton2.size = '600,40'
    $Ex2013ManGroupBoxRadioButton2.Checked = $false
    $Ex2013ManGroupBoxRadioButton2.Text = 'User Mailbox - Out Of Office management.'
    $Ex2013ManGroupBoxRadioButton3 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton3.Location = '20,110'
    $Ex2013ManGroupBoxRadioButton3.size = '600,40'
    $Ex2013ManGroupBoxRadioButton3.Checked = $false
    $Ex2013ManGroupBoxRadioButton3.Text = 'Shared Calendar - management.'
    $Ex2013ManGroupBoxRadioButton4 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton4.Location = '20,150'
    $Ex2013ManGroupBoxRadioButton4.size = '600,40'
    $Ex2013ManGroupBoxRadioButton4.Checked = $false
    $Ex2013ManGroupBoxRadioButton4.Text = 'Distribution List -  management.'
    $Ex2013ManGroupBoxRadioButton5 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton5.Location = '20,190'
    $Ex2013ManGroupBoxRadioButton5.size = '600,40'
    $Ex2013ManGroupBoxRadioButton5.Checked = $false
    $Ex2013ManGroupBoxRadioButton5.Text = 'Disabled User Account - mailbox management.'
    $Ex2013ManGroupBoxRadioButton6 = New-Object System.Windows.Forms.RadioButton
    $Ex2013ManGroupBoxRadioButton6.Location = '20,230'
    $Ex2013ManGroupBoxRadioButton6.size = '600,40'
    $Ex2013ManGroupBoxRadioButton6.Checked = $false
    $Ex2013ManGroupBoxRadioButton6.Text = 'Mobile Device Reset.'
    # Add an OK button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $Ex2013ManForm.Close()
            $Ex2013ManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $Ex2013ManForm.Controls.AddRange(@($Ex2013ManGroupBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $Ex2013ManGroupBox.Controls.AddRange(@($Ex2013ManGroupBoxRadioButton1, $Ex2013ManGroupBoxRadioButton2, $Ex2013ManGroupBoxRadioButton3, $Ex2013ManGroupBoxRadioButton4, $Ex2013ManGroupBoxRadioButton5, $Ex2013ManGroupBoxRadioButton6))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $Ex2013ManForm.AcceptButton = $OKButton
    $Ex2013ManForm.CancelButton = $CancelButton
    # Activate the form
    $Ex2013ManForm.Add_Shown( { $Ex2013ManForm.Activate() })    
    # Get the results from the button click
    $Result = $Ex2013ManForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        # Check the current state of each radio button and respond accordingly
        if ($Ex2013ManGroupBoxRadioButton1.Checked) {
            SharedMailboxManagementForm
        }
        elseif ($Ex2013ManGroupBoxRadioButton2.Checked) {
            OutOfficeManagementForm
        }
        elseif ($Ex2013ManGroupBoxRadioButton3.Checked) {
            SharedCalendarManagementForm
        }
        elseif ($Ex2013ManGroupBoxRadioButton4.Checked) {
            DistributionListManagementForm
        }
        elseif ($Ex2013ManGroupBoxRadioButton5.Checked) {
            DisabledUserManagementForm
        }
        elseif ($Ex2013ManGroupBoxRadioButton6.Checked = $True) {
            MobileResetform
        }
    }
}
Return MainForm
#########################################################################################################
######      Completed -    Create the 'Exchange 2103 User Management main form' form       ############## 
#########################################################################################################
#endregion Mainform