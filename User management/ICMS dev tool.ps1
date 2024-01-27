$UserNameList = Get-ADUser -Filter * -SearchBase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$groups = Get-ADGroup -Filter "name -like 'ICMS Development*' -and name -notlike '*Judiciary'-and name -notlike '*Team' -and name -notlike '*System Admin' -and name -notlike '*Training Learning and Development*' -and name -notlike '*Helpdesk' -and name -notlike '*Finance Manager'" | Select-Object Name | Select-Object -ExpandProperty Name
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\Add User.ico'
$Version = '1.1'
$WinTitle = "ICMS Dev control script v$version."
##############

Function Mainform {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.width = 745
    $MainForm.height = 475
    $MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $MainForm.Controlbox = $false
    $MainForm.Icon = $Icon
    $MainForm.FormBorderStyle = 'Fixed3D'
    $MainForm.Text = $WinTitle
    $MainForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ### Create group 1 box in form. ####
    $MainBox1 = New-Object System.Windows.Forms.GroupBox
    $MainBox1.Location = '40,20'
    $MainBox1.size = '650,125'
    $MainBox1.text = '1. Select a UserName from the dropdown lists:'
    ### Create group 1 box text labels. ###
    $MaintextLabel1 = New-Object System.Windows.Forms.Label
    $MaintextLabel1.Location = '20,40'
    $MaintextLabel1.size = '100,40'
    $MaintextLabel1.Text = 'UserName:' 
    ### Create group 1 box combo boxes. ###
    $MainUserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $MainUserNameComboBox1.Location = '275,35'
    $MainUserNameComboBox1.Size = '350, 310'
    $MainUserNameComboBox1.AutoCompleteMode = 'Suggest'
    $MainUserNameComboBox1.AutoCompleteSource = 'ListItems'
    $MainUserNameComboBox1.Sorted = $true;
    $MainUserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $MainUserNameComboBox1.DataSource = $UsernameList
    $MainUserNameComboBox1.add_SelectedIndexChanged( { $Selected.Text = "$($MainUserNameComboBox1.SelectedItem.ToString())" })
    ### Create group 2 box in form. ###
    $MainBox2 = New-Object System.Windows.Forms.GroupBox
    $MainBox2.Location = '40,170'
    $MainBox2.size = '650,125'
    $MainBox2.text = '2. Add or remove:'
    # Create group 2 box text labels.
    $BoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton1.Location = '20,20'
    $BoxRadioButton1.size = '200,40'
    $BoxRadioButton1.Checked = $true 
    $BoxRadioButton1.Text = 'Add'
    $BoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton2.Location = '20,50'
    $BoxRadioButton2.size = '200,40'
    $BoxRadioButton2.Checked = $false
    $BoxRadioButton2.Text = 'Remove'
    $BoxRadioButton2.Add_Click( {
            $BoxRadioButton1.Checked = $false
            $BoxRadioButton2.Checked = $true })
    $BoxRadioButton1.Add_Click( {
            $BoxRadioButton1.Checked = $true
            $BoxRadioButton2.Checked = $false })
    ### Create group 3 box in form. ###
    $MainBox3 = New-Object System.Windows.Forms.GroupBox
    $MainBox3.Location = '40,320'
    $MainBox3.size = '650,50'
    $MainBox3.text = '3. Click Ok to add or remove permissions or Cancel:'
    $MainBox3.button
    $Selected = New-Object System.Windows.Forms.Label
    $Selected.Location = '20,20'
    $Selected.Size = '200,20'
    $Selected.ForeColor = 'Blue'
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
    $CancelButton.Text = 'Cancel'
    $CancelButton.add_Click( {
            $MainForm.Close()
            $MainForm.Dispose() })
    ### Add all the Form controls ### 
    $MainForm.Controls.AddRange(@($MainBox1, $MainBox2, $MainBox3, $OKButton, $CancelButton))
    $MainBox1.Controls.AddRange(@($MaintextLabel1, $MainUserNameComboBox1))
    $MainBox2.Controls.AddRange(@($BoxRadioButton1, $BoxRadioButton2))
    $MainBox3.Controls.AddRange(@($Selected))
    #### Assign the Accept and Cancel options in the form ### 
    $MainForm.AcceptButton = $OKButton
    $MainForm.CancelButton = $CancelButton
    #### Activate the form ###
    $MainForm.Add_Shown( { $MainForm.Activate() })    
    #### Get the results from the button click ###
    $dialogResult = $MainForm.ShowDialog()
    # If the OK button is selected
    if ($dialogResult -eq 'OK') {
        $AddUser = $Selected.Text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$AddUser'" | Select-Object -ExpandProperty 'SamAccountName'
        if ($BoxRadioButton1.Checked) {
            ForEach ($group in $groups) {
                Add-ADGroupMember -Identity $group -Members $UserSamAccountName
                Write-Output "$UserSamAccountName added to $group"
            }
        }
        else {
            ForEach ($group in $groups) {
                remove-ADGroupMember -Identity $group -Members $UserSamAccountName -Confirm:$false
                Write-Output "$UserSamAccountName Removed to $group" 
            }
            Write-Output -InputObject "Complete. Press any key to continue..."
            [void]$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}
Mainform 