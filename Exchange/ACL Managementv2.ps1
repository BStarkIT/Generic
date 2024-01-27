# ACL Management v1.00
# Author        Brian Stark
# Date          20/04/2020
# Version       1.00
# Purpose       To manage ACL groups
# Useage        Easier mantance of ACL groups within the current ACL guidelines
# 
# Revisions     V1.00 20/04/2020 BS - Basic outline & basic functions
#
## Layout of file:
#
# Main Form 
#
#######################################################################################################
###                 Set icon for all forms and subforms                                             ###
#######################################################################################################
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\acl.ico'
#######################################################################################################
###                 Get ACL's from AD                                                               ###
#######################################################################################################
$ACLlist = Get-ADGroup -Filter "name -like 'ACL*'" | Select-Object Name | Select-Object -ExpandProperty Name
########################################################################################################
###                 Get listof UserNames from AD OU's                                               ####
########################################################################################################
$Users1 = Get-aduser –filter * -searchbase 'ou=tribunalusers,ou=tribunals,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users2 = Get-aduser –filter * -searchbase 'ou=sheriffsparttime,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users3 = Get-aduser –filter * -searchbase 'ou=scs employees,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Where-Object { ($_.DistinguishedName -notlike '*OU=deleted users,*') -and ($_.DistinguishedName -notlike '*OU=it administrators,*') } | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users4 = Get-aduser –filter * -searchbase 'ou=JP,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users5 = Get-aduser –filter * -searchbase 'ou=judges,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users6 = Get-aduser –filter * -searchbase 'ou=sheriffs,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users7 = Get-aduser –filter * -searchbase 'ou=sheriffsprincipal,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users8 = Get-aduser –filter * -searchbase 'ou=sheriffssummary,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users9 = Get-aduser –filter * -searchbase 'ou=sheriffsretired,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users10 = Get-aduser –filter * -searchbase 'ou=courts,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users11 = Get-aduser –filter * -searchbase 'ou=judiciary,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users12 = Get-aduser –filter * -searchbase 'ou=tribunals,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users13 = Get-aduser –filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
$UserNameList = $Users1 + $users2 + $users3 + $users4 + $users5 + $users6 + $users7 + $users8 + $users9 + $users10 + $users11 + $users12 + $users13
#######################################################################################################
###
#####################################################################
#       Create SubForm 'ACL List - List current members of a List   #
#####################################################################
Function ListACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $ACLListListForm = New-Object System.Windows.Forms.Form
    $ACLListListForm.width = 745
    $ACLListListForm.height = 475
    $ACLListListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $ACLListListForm.Controlbox = $false
    $ACLListListForm.Icon = $Icon
    $ACLListListForm.FormBorderStyle = 'Fixed3D'
    $ACLListListForm.Text = 'ACL - List current members of a List.'
    $ACLListListForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $ACLListListBox1 = New-Object System.Windows.Forms.GroupBox
    $ACLListListBox1.Location = '40,20'
    $ACLListListBox1.size = '650,125'
    $ACLListListBox1.text = '1. Select an ACL List from the dropdown list:'
    ### Create group 1 box text labels. ###
    $ACLListListtextLabel2 = New-Object System.Windows.Forms.Label
    $ACLListListtextLabel2.Location = '20,50'
    $ACLListListtextLabel2.size = '200,40'
    $ACLListListtextLabel2.Text = 'ACL List:' 
    ### Create group 1 box combo boxes. ###
    $ACLListListMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $ACLListListMBNameComboBox2.Location = '275,45'
    $ACLListListMBNameComboBox2.Size = '350, 350'
    $ACLListListMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $ACLListListMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $ACLListListMBNameComboBox2.Sorted = $true;
    $ACLListListMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $ACLListListMBNameComboBox2.DataSource = $ACLlist
    $ACLListListMBNameComboBox2.Add_SelectedIndexChanged( { $ACLListListSelectedNametextLabel6.Text = "$($ACLListListMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $ACLListListBox2 = New-Object System.Windows.Forms.GroupBox
    $ACLListListBox2.Location = '40,170'
    $ACLListListBox2.size = '650,125'
    $ACLListListBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $ACLListListtextLabel3 = New-Object System.Windows.Forms.Label
    $ACLListListtextLabel3.Location = '40,40'
    $ACLListListtextLabel3.size = '400,40'
    $ACLListListtextLabel3.Text = 'List current members of ACL List:' 
    $ACLListListSelectedNametextLabel6 = New-Object System.Windows.Forms.Label
    $ACLListListSelectedNametextLabel6.Location = '100,80'
    $ACLListListSelectedNametextLabel6.Size = '400,40'
    $ACLListListSelectedNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $ACLListListBox3 = New-Object System.Windows.Forms.GroupBox
    $ACLListListBox3.Location = '40,320'
    $ACLListListBox3.size = '650,30'
    $ACLListListBox3.text = '3. Click Ok to List current members or Cancel:'
    $ACLListListBox3.button
    ### Add an OK button ###
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '590,370'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '470,370'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.Add_Click( {
            $ACLListListForm.Close()
            $ACLListListForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $ACLListListForm.Controls.AddRange(@($ACLListListBox1, $ACLListListBox2, $ACLListListBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $ACLListListBox1.Controls.AddRange(@($ACLListListtextLabel2, $ACLListListMBNameComboBox2))
    $ACLListListBox2.Controls.AddRange(@($ACLListListtextLabel3, $DistListListtextLabel5, $ACLListListSelectedNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $ACLListListForm.AcceptButton = $OKButton
    $ACLListListForm.CancelButton = $CancelButton
    #### Activate the form ###
    $ACLListListForm.Add_Shown( { $ACLListListForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $ACLListListForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################
        #               Don't accept null ACL               # 
        #####################################################
        if ($ACLListListSelectedNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL list !!!!!  Trying to enter blank fields is never a good idea.", 'ACL List - List current members of a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $ACLListListForm.Close()
            $ACLListListForm.Dispose()
            break
        }
        #####################################################
        #               Get List of Names                   # 
        #####################################################
        $ACLListName = $ACLListListSelectedNametextLabel6.Text 
        Get-adGroupMember -Identity $ACLListName | Select-Object Name | Out-GridView -Title "List of current members of the $ACLListName ACL list" -Wait
        $ACLListListForm.Close()
        $ACLListListForm.Dispose()
        Return ListACLManagementForm
    }
}
#############################################################################
#   Completed -  Create SubForm 'ACL List - List current members of a List' #
#############################################################################
#
######################################################################
####  Create SubForm Add - Add user to ACL   ###
######################################################################
Function AddtoACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddACLFullForm = New-Object System.Windows.Forms.Form
    $AddACLFullForm.width = 780
    $AddACLFullForm.height = 500
    $AddACLFullForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddACLFullForm.Controlbox = $false
    $AddACLFullForm.Icon = $Icon
    $AddACLFullForm.FormBorderStyle = 'Fixed3D'
    $AddACLFullForm.Text = 'ACL - Add User to ACL.'
    $AddACLFullForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddACLFullBox1 = New-Object System.Windows.Forms.GroupBox
    $AddACLFullBox1.Location = '40,40'
    $AddACLFullBox1.size = '700,125'
    $AddACLFullBox1.text = '1. Select a UserName and ACL from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $AddACLFulltextLabel1 = New-Object System.Windows.Forms.Label
    $AddACLFulltextLabel1.Location = '20,40'
    $AddACLFulltextLabel1.size = '150,40'
    $AddACLFulltextLabel1.Text = 'UserName:' 
    $AddACLFulltextLabel2 = New-Object System.Windows.Forms.Label
    $AddACLFulltextLabel2.Location = '20,80'
    $AddACLFulltextLabel2.size = '150,40'
    $AddACLFulltextLabel2.Text = 'ACL:' 
    ### Create group 1 box combo boxes. ###
    $AddACLFullUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $AddACLFullUserNameComboBox1.Location = '325,35'
    $AddACLFullUserNameComboBox1.Size = '350, 310'
    $AddACLFullUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $AddACLFullUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $AddACLFullUserNameComboBox1.Sorted = $true;
    $AddACLFullUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddACLFullUserNameComboBox1.DataSource = $UsernameList
    $AddACLFullUserNameComboBox1.add_SelectedIndexChanged( { $AddACLFullSelectedUserNametextLabel4.Text = "$($AddACLFullUserNameComboBox1.SelectedItem.ToString())" })
    $AddACLFullACLNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddACLFullACLNameComboBox2.Location = '325,75'
    $AddACLFullACLNameComboBox2.Size = '350, 350'
    $AddACLFullACLNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddACLFullACLNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddACLFullACLNameComboBox2.Sorted = $true;
    $AddACLFullACLNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddACLFullACLNameComboBox2.DataSource = $ACLlist 
    $AddACLFullACLNameComboBox2.add_SelectedIndexChanged( { $AddACLFullSelectedACLNametextLabel6.Text = "$($AddACLFullACLNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddACLFullBox2 = New-Object System.Windows.Forms.GroupBox
    $AddACLFullBox2.Location = '40,190'
    $AddACLFullBox2.size = '700,125'
    $AddACLFullBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $AddACLFulltextLabel3 = New-Object System.Windows.Forms.Label
    $AddACLFulltextLabel3.Location = '40,40'
    $AddACLFulltextLabel3.size = '100,40'
    $AddACLFulltextLabel3.Text = 'The User:' 
    $AddACLFullSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $AddACLFullSelectedUserNametextLabel4.Location = '30,80'
    $AddACLFullSelectedUserNametextLabel4.Size = '200,40'
    $AddACLFullSelectedUserNametextLabel4.ForeColor = 'Blue'
    $AddACLFulltextLabel5 = New-Object System.Windows.Forms.Label
    $AddACLFulltextLabel5.Location = '275,40'
    $AddACLFulltextLabel5.size = '400,40'
    $AddACLFulltextLabel5.Text = 'Will be added to the ACL:'
    $AddACLFullSelectedACLNametextLabel6 = New-Object System.Windows.Forms.Label
    $AddACLFullSelectedACLNametextLabel6.Location = '350,80'
    $AddACLFullSelectedACLNametextLabel6.Size = '200,40'
    $AddACLFullSelectedACLNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $AddACLFullBox3 = New-Object System.Windows.Forms.GroupBox
    $AddACLFullBox3.Location = '40,340'
    $AddACLFullBox3.size = '700,30'
    $AddACLFullBox3.text = '3. Click Ok to add ACL permissions or Cancel:'
    $AddACLFullBox3.button
    ### Add an OK button ###
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $AddACLFullForm.Close()
            $AddACLFullForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $AddACLFullForm.Controls.AddRange(@($AddACLFullBox1, $AddACLFullBox2, $AddACLFullBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $AddACLFullBox1.Controls.AddRange(@($AddACLFulltextLabel1, $AddACLFulltextLabel2, $AddACLFullUserNameComboBox1, $AddACLFullACLNameComboBox2))
    $AddACLFullBox2.Controls.AddRange(@($AddACLFulltextLabel3, $AddACLFullSelectedUserNametextLabel4, $AddACLFulltextLabel5, $AddACLFullSelectedACLNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $AddACLFullForm.AcceptButton = $OKButton
    $AddACLFullForm.CancelButton = $CancelButton
    #### Activate the form ###
    $AddACLFullForm.Add_Shown( { $AddACLFullForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $AddACLFullForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or ACL  ################ 
        #####################################################################
        if ($AddACLFullSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'ACL - Add User to ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddACLFullForm.Close()
            $AddACLFullForm.Dispose()
            break
        }
        Elseif ($AddACLFullSelectedACLNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL !!!!!  Trying to enter blank fields is never a good idea.", 'ACL - Add User to ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddACLFullForm.Close()
            $AddACLFullForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get ACL Name:  ####
        #####################################################################
        $AddUser = $AddACLFullSelectedUserNametextLabel4.Text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$AddUser'" | Select-Object -ExpandProperty 'SamAccountName'
        $ACLListName = $AddACLFullSelectedACLNametextLabel6.Text 
        #####################################################################
        ######                Add user        ###############
        #####################################################################
        Add-ADGroupMember –Identity $ACLListName -Members $UserSamAccountName
        #####################################################################
        ##################  Message complete message  #######################
        #####################################################################
        Add-Type -AssemblyName System.Windows.Forms 
        [System.Windows.Forms.MessageBox]::Show("The user ( $($AddACLFullUserNameComboBox1.SelectedItem.ToString()) )`nhas been added to the ( $($AddACLFullACLNameComboBox2.SelectedItem.ToString()) ) ACL.", 'ACL - Add User to ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $AddACLFullForm.Close()
        $AddACLFullForm.Dispose()
        Return AddtoACLManagementForm
        
    }
}
#########################################################################
##  Completed -Create SubForm Add - Add user to ACL ##
#########################################################################
#
######################################################################
####  Create SubForm Remove - remove user from ACL   ###
######################################################################
Function RemovefromACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $RmoveACL1FullForm = New-Object System.Windows.Forms.Form
    $RmoveACL1FullForm.width = 780
    $RmoveACL1FullForm.height = 500
    $RmoveACL1FullForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RmoveACL1FullForm.Controlbox = $false
    $RmoveACL1FullForm.Icon = $Icon
    $RmoveACL1FullForm.FormBorderStyle = 'Fixed3D'
    $RmoveACL1FullForm.Text = 'ACL - Remove User from ACL.'
    $RmoveACL1FullForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $RmoveACL1FullBox1 = New-Object System.Windows.Forms.GroupBox
    $RmoveACL1FullBox1.Location = '40,40'
    $RmoveACL1FullBox1.size = '700,125'
    $RmoveACL1FullBox1.text = '1. Select a UserName and ACL from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $RmoveACL1FulltextLabel1 = New-Object System.Windows.Forms.Label
    $RmoveACL1FulltextLabel1.Location = '20,40'
    $RmoveACL1FulltextLabel1.size = '150,40'
    $RmoveACL1FulltextLabel1.Text = 'UserName:' 
    $RmoveACL1FulltextLabel2 = New-Object System.Windows.Forms.Label
    $RmoveACL1FulltextLabel2.Location = '20,80'
    $RmoveACL1FulltextLabel2.size = '150,40'
    $RmoveACL1FulltextLabel2.Text = 'ACL:' 
    ### Create group 1 box combo boxes. ###
    $RmoveACL1FullUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $RmoveACL1FullUserNameComboBox1.Location = '325,35'
    $RmoveACL1FullUserNameComboBox1.Size = '350, 310'
    $RmoveACL1FullUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $RmoveACL1FullUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $RmoveACL1FullUserNameComboBox1.Sorted = $true;
    $RmoveACL1FullUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RmoveACL1FullUserNameComboBox1.DataSource = $UsernameList
    $RmoveACL1FullUserNameComboBox1.add_SelectedIndexChanged( { $RmoveACL1FullSelectedUserNametextLabel4.Text = "$($RmoveACL1FullUserNameComboBox1.SelectedItem.ToString())" })
    $RmoveACL1FullACLNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $RmoveACL1FullACLNameComboBox2.Location = '325,75'
    $RmoveACL1FullACLNameComboBox2.Size = '350, 350'
    $RmoveACL1FullACLNameComboBox2.AutoCompleteMode = 'Suggest'
    $RmoveACL1FullACLNameComboBox2.AutoCompleteSource = 'ListItems'
    $RmoveACL1FullACLNameComboBox2.Sorted = $true;
    $RmoveACL1FullACLNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RmoveACL1FullACLNameComboBox2.DataSource = $ACLlist 
    $RmoveACL1FullACLNameComboBox2.add_SelectedIndexChanged( { $RmoveACL1FullSelectedACLNametextLabel6.Text = "$($RmoveACL1FullACLNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $RmoveACL1FullBox2 = New-Object System.Windows.Forms.GroupBox
    $RmoveACL1FullBox2.Location = '40,190'
    $RmoveACL1FullBox2.size = '700,125'
    $RmoveACL1FullBox2.text = '2. Check the details below are correct before proceeding:'
    # Create group 2 box text labels.
    $RmoveACL1FulltextLabel3 = New-Object System.Windows.Forms.Label
    $RmoveACL1FulltextLabel3.Location = '40,40'
    $RmoveACL1FulltextLabel3.size = '100,40'
    $RmoveACL1FulltextLabel3.Text = 'The User:' 
    $RmoveACL1FullSelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $RmoveACL1FullSelectedUserNametextLabel4.Location = '30,80'
    $RmoveACL1FullSelectedUserNametextLabel4.Size = '200,40'
    $RmoveACL1FullSelectedUserNametextLabel4.ForeColor = 'Blue'
    $RmoveACL1FulltextLabel5 = New-Object System.Windows.Forms.Label
    $RmoveACL1FulltextLabel5.Location = '275,40'
    $RmoveACL1FulltextLabel5.size = '400,40'
    $RmoveACL1FulltextLabel5.Text = 'Will be removed from the ACL:'
    $RmoveACL1FullSelectedACLNametextLabel6 = New-Object System.Windows.Forms.Label
    $RmoveACL1FullSelectedACLNametextLabel6.Location = '350,80'
    $RmoveACL1FullSelectedACLNametextLabel6.Size = '200,40'
    $RmoveACL1FullSelectedACLNametextLabel6.ForeColor = 'Blue'
    ### Create group 3 box in form. ###
    $RmoveACL1FullBox3 = New-Object System.Windows.Forms.GroupBox
    $RmoveACL1FullBox3.Location = '40,340'
    $RmoveACL1FullBox3.size = '700,30'
    $RmoveACL1FullBox3.text = '3. Click Ok to remove ACL permissions or Cancel:'
    $RmoveACL1FullBox3.button
    ### Add an OK button ###
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $RmoveACL1FullForm.Close()
            $RmoveACL1FullForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $RmoveACL1FullForm.Controls.AddRange(@($RmoveACL1FullBox1, $RmoveACL1FullBox2, $RmoveACL1FullBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $RmoveACL1FullBox1.Controls.AddRange(@($RmoveACL1FulltextLabel1, $RmoveACL1FulltextLabel2, $RmoveACL1FullUserNameComboBox1, $RmoveACL1FullACLNameComboBox2))
    $RmoveACL1FullBox2.Controls.AddRange(@($RmoveACL1FulltextLabel3, $RmoveACL1FullSelectedUserNametextLabel4, $RmoveACL1FulltextLabel5, $RmoveACL1FullSelectedACLNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $RmoveACL1FullForm.AcceptButton = $OKButton
    $RmoveACL1FullForm.CancelButton = $CancelButton
    #### Activate the form ###
    $RmoveACL1FullForm.Add_Shown( { $RmoveACL1FullForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $RmoveACL1FullForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################################
        ########   Don't accept null username or ACL  ################ 
        #####################################################################
        if ($RmoveACL1FullSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'ACL - Remove User from ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RmoveACL1FullForm.Close()
            $RmoveACL1FullForm.Dispose()
            break
        }
        Elseif ($RmoveACL1FullSelectedACLNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL !!!!!  Trying to enter blank fields is never a good idea.", 'ACL - Remove User from ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $RmoveACL1FullForm.Close()
            $RmoveACL1FullForm.Dispose()
            break
        }
        #####################################################################
        ##########  get user samaccountname from user name:  ################ 
        ###  get ACL Name:  ####
        #####################################################################
        $RemoveUser = $RmoveACL1FullSelectedUserNametextLabel4.Text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$RemoveUser'"| Select-Object -ExpandProperty 'SamAccountName'
        $ACLListName = $RmoveACL1FullSelectedACLNametextLabel6.Text 
        #####################################################################
        ######                Add user        ###############
        #####################################################################
        remove-ADGroupMember –Identity $ACLListName -Members $UserSamAccountName
        #####################################################################
        ##################  Message complete message  #######################
        #####################################################################
        Add-Type -AssemblyName System.Windows.Forms 
        [System.Windows.Forms.MessageBox]::Show("The user ( $($RmoveACL1FullUserNameComboBox1.SelectedItem.ToString()) )`nhas been removed from the ( $($RmoveACL1FullACLNameComboBox2.SelectedItem.ToString()) ) ACL.", 'ACL - Remove User from ACL.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $RmoveACL1FullForm.Close()
        $RmoveACL1FullForm.Dispose()
        Return RemovefromACLManagementForm
        
    }
}
#########################################################################
##  Completed -Create SubForm Remove - remove user from ACL ##
#########################################################################
#
#
#######################################################################################################
###                 Create the 'ACL Management main form' form                                      ### 
#######################################################################################################
Function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details size of your form
    $ACLManForm = New-Object System.Windows.Forms.Form
    $ACLManForm.width = 780
    $ACLManForm.height = 350
    $ACLManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $ACLManForm.MinimizeBox = $False
    $ACLManForm.MaximizeBox = $False
    $ACLManForm.FormBorderStyle = 'Fixed3D'
    $ACLManForm.Text = 'ACL Management Main Form v1.00'
    $ACLManForm.Icon = $Icon
    $ACLManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    # Create a group that will contain your radio buttons
    $ACLManGroupBox = New-Object System.Windows.Forms.GroupBox
    $ACLManGroupBox.Location = '40,50'
    $ACLManGroupBox.size = '700,160'
    $ACLManGroupBox.text = 'Select an option.'
    # Create the collection of radio buttons
    $ACLManGroupBoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton1.Location = '20,30'
    $ACLManGroupBoxRadioButton1.size = '600,40'
    $ACLManGroupBoxRadioButton1.Checked = $false
    $ACLManGroupBoxRadioButton1.Text = 'View Staff in an ACL.'
    $ACLManGroupBoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton2.Location = '20,70'
    $ACLManGroupBoxRadioButton2.size = '600,40'
    $ACLManGroupBoxRadioButton2.Checked = $false
    $ACLManGroupBoxRadioButton2.Text = 'Add Staff to an ACL.'
    $ACLManGroupBoxRadioButton3 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton3.Location = '20,110'
    $ACLManGroupBoxRadioButton3.size = '600,40'
    $ACLManGroupBoxRadioButton3.Checked = $false
    $ACLManGroupBoxRadioButton3.Text = 'Remove Staff from an ACL.'
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '640,250'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,250'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $ACLManForm.Close()
            $ACLManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $ACLManForm.Controls.AddRange(@($ACLManGroupBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $ACLManGroupBox.Controls.AddRange(@($ACLManGroupBoxRadioButton1, $ACLManGroupBoxRadioButton2, $ACLManGroupBoxRadioButton3))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $ACLManForm.AcceptButton = $OKButton
    $ACLManForm.CancelButton = $CancelButton
    # Activate the form
    $ACLManForm.Add_Shown( { $ACLManForm.Activate() })    
    # Get the results from the button click
    $Result = $ACLManForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        # Check the current state of each radio button and respond accordingly
        if ($ACLManGroupBoxRadioButton1.Checked) {
            ListACLManagementForm
        }
        elseif ($ACLManGroupBoxRadioButton2.Checked) {
            AddtoACLManagementForm
        }
        elseif ($ACLManGroupBoxRadioButton3.Checked = $True) {
            RemovefromACLManagementForm
        }
    }
}
Return MainForm
#######################################################################################################
###                 Completed - Create ACL Management main form' form                               ### 
#######################################################################################################