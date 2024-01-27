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
# Build from here
#####################################################################
#       Create    #
#####################################################################

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
#####################################################################
#       Create SubForm 'ACL ADD - Add members of to a ACL List'     #
#####################################################################
Function AddtoACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddtoACLListForm = New-Object System.Windows.Forms.Form
    $AddtoACLListForm.width = 745
    $AddtoACLListForm.height = 475
    $AddtoACLListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddtoACLListForm.Controlbox = $false
    $AddtoACLListForm.Icon = $Icon
    $AddtoACLListForm.FormBorderStyle = 'Fixed3D'
    $AddtoACLListForm.Text = 'ACL - Add members to a list.'
    $AddtoACLListForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddtoACLListBox1 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox1.Location = '40,20'
    $AddtoACLListBox1.size = '650,125'
    $AddtoACLListBox1.text = '1. Select an ACL List from the dropdown list:'
    ### Create group 1 box text labels. ###
    $AddtoACLListtextLabel1 = New-Object System.Windows.Forms.Label
    $AddtoACLListtextLabel1.Location = '20,50'
    $AddtoACLListtextLabel1.size = '200,40'
    $AddtoACLListtextLabel1.Text = 'ACL List:' 
    $AddtoACLListtextLabel2 = New-Object System.Windows.Forms.Label
    $AddtoACLListtextLabel2.Location = '20,80'
    $AddtoACLListtextLabel2.size = '150,40'
    $AddtoACLListtextLabel2.Text = 'MailBoxName:' 
    ### Create group 1 box combo boxes. ###
    $AddtoACLListMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $AddtoACLListMBNameComboBox2.Location = '275,45'
    $AddtoACLListMBNameComboBox2.Size = '350, 350'
    $AddtoACLListMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddtoACLListMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddtoACLListMBNameComboBox2.Sorted = $true;
    $AddtoACLListMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddtoACLListMBNameComboBox2.DataSource = $ACLlist
    $AddtoACLListMBNameComboBox2.Add_SelectedIndexChanged( { $ACLListListSelectedNametextLabel6.Text = "$($ACLListListMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddtoACLListBox2 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox2.Location = '40,170'
    $AddtoACLListBox2.size = '650,125'
    $AddtoACLListBox2.text = '2. Check the details below are correct before proceeding:'
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
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $CheckPathListForm.Controls.AddRange(@($ACLListListBox1, $ACLListListBox2, $ACLListListBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $ACLListListBox1.Controls.AddRange(@($ACLListListtextLabel2, $ACLListListMBNameComboBox2))
    $ACLListListBox2.Controls.AddRange(@($ACLListListtextLabel3, $DistListListtextLabel5, $ACLListListSelectedNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $CheckPathListForm.AcceptButton = $OKButton
    $CheckPathListForm.CancelButton = $CancelButton
    #### Activate the form ###
    $CheckPathListForm.Add_Shown( { $CheckPathListForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckPathListForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################
        #               Don't accept null ACL               # 
        #####################################################
        if ($ACLListListSelectedNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL list !!!!!  Trying to enter blank fields is never a good idea.", 'ACL List - List current members of a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            break
        }
        #####################################################
        #               Get List of Names                   # 
        #####################################################
        $ACLListName = $ACLListListSelectedNametextLabel6.Text 
        Get-adGroupMember -Identity $ACLListName | Select-Object Name | Out-GridView -Title "List of current members of the $ACLListName ACL list" -Wait
        $CheckPathListForm.Close()
        $CheckPathListForm.Dispose()
        Return AddtoACLManagementForm
    }
}
#############################################################################
#   Completed -  Create SubForm 'ACL List - List current members of a List  #
#############################################################################
#
#####################################################################
#       Create SubForm 'ACL Remove - remove current member from a List'     #
#####################################################################
Function RemovefromACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddtoACLListForm = New-Object System.Windows.Forms.Form
    $AddtoACLListForm.width = 745
    $AddtoACLListForm.height = 475
    $AddtoACLListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddtoACLListForm.Controlbox = $false
    $AddtoACLListForm.Icon = $Icon
    $AddtoACLListForm.FormBorderStyle = 'Fixed3D'
    $AddtoACLListForm.Text = 'ACL - List current members of a List.'
    $AddtoACLListForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddtoACLListBox1 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox1.Location = '40,20'
    $AddtoACLListBox1.size = '650,125'
    $AddtoACLListBox1.text = '1. Select an ACL List from the dropdown list:'
    ### Create group 1 box text labels. ###
    $AddtoACLListtextLabel2 = New-Object System.Windows.Forms.Label
    $AddtoACLListtextLabel2.Location = '20,50'
    $AddtoACLListtextLabel2.size = '200,40'
    $AddtoACLListtextLabel2.Text = 'ACL List:' 
    ### Create group 1 box combo boxes. ###
    $AddtoACLListMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $ACLListListMBNameComboBox2.Location = '275,45'
    $AddtoACLListMBNameComboBox2.Size = '350, 350'
    $AddtoACLListMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddtoACLListMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddtoACLListMBNameComboBox2.Sorted = $true;
    $AddtoACLListMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddtoACLListMBNameComboBox2.DataSource = $ACLlist
    $AddtoACLListMBNameComboBox2.Add_SelectedIndexChanged( { $ACLListListSelectedNametextLabel6.Text = "$($ACLListListMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddtoACLListBox2 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox2.Location = '40,170'
    $AddtoACLListBox2.size = '650,125'
    $AddtoACLListBox2.text = '2. Check the details below are correct before proceeding:'
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
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $CheckPathListForm.Controls.AddRange(@($ACLListListBox1, $ACLListListBox2, $ACLListListBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $ACLListListBox1.Controls.AddRange(@($ACLListListtextLabel2, $ACLListListMBNameComboBox2))
    $ACLListListBox2.Controls.AddRange(@($ACLListListtextLabel3, $DistListListtextLabel5, $ACLListListSelectedNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $CheckPathListForm.AcceptButton = $OKButton
    $CheckPathListForm.CancelButton = $CancelButton
    #### Activate the form ###
    $CheckPathListForm.Add_Shown( { $CheckPathListForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckPathListForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################
        #               Don't accept null ACL               # 
        #####################################################
        if ($ACLListListSelectedNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL list !!!!!  Trying to enter blank fields is never a good idea.", 'ACL List - List current members of a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            break
        }
        #####################################################
        #               Get List of Names                   # 
        #####################################################
        $ACLListName = $ACLListListSelectedNametextLabel6.Text 
        Get-adGroupMember -Identity $ACLListName | Select-Object Name | Out-GridView -Title "List of current members of the $ACLListName ACL list" -Wait
        $CheckPathListForm.Close()
        $CheckPathListForm.Dispose()
        Return RemovefromACLManagementForm
    }
}
#############################################################################
# Completed -  Create SubForm 'ACL Remove - remove current member from a List  #
#############################################################################
#
#####################################################################
#       Create SubForm 'Create ACL'     #
#####################################################################
Function CreateACLManagementForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AddtoACLListForm = New-Object System.Windows.Forms.Form
    $AddtoACLListForm.width = 745
    $AddtoACLListForm.height = 475
    $AddtoACLListForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AddtoACLListForm.Controlbox = $false
    $AddtoACLListForm.Icon = $Icon
    $AddtoACLListForm.FormBorderStyle = 'Fixed3D'
    $AddtoACLListForm.Text = 'ACL - List current members of a List.'
    $AddtoACLListForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $AddtoACLListBox1 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox1.Location = '40,20'
    $AddtoACLListBox1.size = '650,125'
    $AddtoACLListBox1.text = '1. Select an ACL List from the dropdown list:'
    ### Create group 1 box text labels. ###
    $AddtoACLListtextLabel2 = New-Object System.Windows.Forms.Label
    $AddtoACLListtextLabel2.Location = '20,50'
    $AddtoACLListtextLabel2.size = '200,40'
    $AddtoACLListtextLabel2.Text = 'ACL List:' 
    ### Create group 1 box combo boxes. ###
    $AddtoACLListMBNameComboBox2 = New-Object System.Windows.Forms.ComboBox
    $ACLListListMBNameComboBox2.Location = '275,45'
    $AddtoACLListMBNameComboBox2.Size = '350, 350'
    $AddtoACLListMBNameComboBox2.AutoCompleteMode = 'Suggest'
    $AddtoACLListMBNameComboBox2.AutoCompleteSource = 'ListItems'
    $AddtoACLListMBNameComboBox2.Sorted = $true;
    $AddtoACLListMBNameComboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $AddtoACLListMBNameComboBox2.DataSource = $ACLlist
    $AddtoACLListMBNameComboBox2.Add_SelectedIndexChanged( { $ACLListListSelectedNametextLabel6.Text = "$($ACLListListMBNameComboBox2.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $AddtoACLListBox2 = New-Object System.Windows.Forms.GroupBox
    $AddtoACLListBox2.Location = '40,170'
    $AddtoACLListBox2.size = '650,125'
    $AddtoACLListBox2.text = '2. Check the details below are correct before proceeding:'
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
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            Return MainForm })
    ### Add all the Form controls ### 
    $CheckPathListForm.Controls.AddRange(@($ACLListListBox1, $ACLListListBox2, $ACLListListBox3, $OKButton, $CancelButton))
    #### Add all the GroupBox controls ###
    $ACLListListBox1.Controls.AddRange(@($ACLListListtextLabel2, $ACLListListMBNameComboBox2))
    $ACLListListBox2.Controls.AddRange(@($ACLListListtextLabel3, $DistListListtextLabel5, $ACLListListSelectedNametextLabel6))
    #### Assign the Accept and Cancel options in the form ### 
    $CheckPathListForm.AcceptButton = $OKButton
    $CheckPathListForm.CancelButton = $CancelButton
    #### Activate the form ###
    $CheckPathListForm.Add_Shown( { $CheckPathListForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $CheckPathListForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        #####################################################
        #               Don't accept null ACL               # 
        #####################################################
        if ($ACLListListSelectedNametextLabel6.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a ACL list !!!!!  Trying to enter blank fields is never a good idea.", 'ACL List - List current members of a List', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $CheckPathListForm.Close()
            $CheckPathListForm.Dispose()
            break
        }
        #####################################################
        #               Get List of Names                   # 
        #####################################################
        $ACLListName = $ACLListListSelectedNametextLabel6.Text 
        Get-adGroupMember -Identity $ACLListName | Select-Object Name | Out-GridView -Title "List of current members of the $ACLListName ACL list" -Wait
        $CheckPathListForm.Close()
        $CheckPathListForm.Dispose()
        Return CreateACLManagementForm
    }
}
#############################################################################
# Completed -  Create SubForm 'Create ACL'  #
#############################################################################
#
#
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
    $ACLManForm.height = 500
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
    $ACLManGroupBox.size = '700,320'
    $ACLManGroupBox.text = 'Select an option.'
    # Create the collection of radio buttons
    $ACLManGroupBoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton1.Location = '20,70'
    $ACLManGroupBoxRadioButton1.size = '600,40'
    $ACLManGroupBoxRadioButton1.Checked = $false
    $ACLManGroupBoxRadioButton1.Text = 'View Staff in an ACL.'
    $ACLManGroupBoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton2.Location = '20,110'
    $ACLManGroupBoxRadioButton2.size = '600,40'
    $ACLManGroupBoxRadioButton2.Checked = $false
    $ACLManGroupBoxRadioButton2.Text = 'Add Staff to an ACL.'
    $ACLManGroupBoxRadioButton3 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton3.Location = '20,150'
    $ACLManGroupBoxRadioButton3.size = '600,40'
    $ACLManGroupBoxRadioButton3.Checked = $false
    $ACLManGroupBoxRadioButton3.Text = 'Remove Staff from an ACL.'
    $ACLManGroupBoxRadioButton4 = New-Object System.Windows.Forms.RadioButton
    $ACLManGroupBoxRadioButton4.Location = '20,190'
    $ACLManGroupBoxRadioButton4.size = '600,40'
    $ACLManGroupBoxRadioButton4.Checked = $false
    $ACLManGroupBoxRadioButton4.Text = 'Add an ACL.'
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $ACLManForm.Close()
            $ACLManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $ACLManForm.Controls.AddRange(@($ACLManGroupBox, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $ACLManGroupBox.Controls.AddRange(@($ACLManGroupBoxRadioButton1, $ACLManGroupBoxRadioButton2, $ACLManGroupBoxRadioButton3, $ACLManGroupBoxRadioButton4))
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
        elseif ($ACLManGroupBoxRadioButton3.Checked) {
            RemovefromACLManagementForm
        }
        elseif ($ACLManGroupBoxRadioButton4.Checked = $True) {
            CreateACLManagementForm
        }
    }
}
Return MainForm
#######################################################################################################
###                 Completed - Create ACL Management main form' form                               ### 
#######################################################################################################