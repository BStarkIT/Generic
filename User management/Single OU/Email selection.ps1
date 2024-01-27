# Start of script:
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\IT\Enterprise Team\UserManagement\Logging\Email\$Date.txt" -append
$Version = '3.00'
$WinTitle = "Email Selection script v$version."
# Set icon for all forms and subforms
#
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\User.ico'
#
# Get listof UserNames from AD OU's
#
$UserNameList = Get-ADUser -filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
#
#
function EmailForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $RenameNameForm = New-Object System.Windows.Forms.Form
    $RenameNameForm.width = 930
    $RenameNameForm.height = 550
    $RenameNameForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RenameNameForm.Controlbox = $false
    $RenameNameForm.Icon = $Icon
    $RenameNameForm.FormBorderStyle = 'Fixed3D'
    $RenameNameForm.Text = 'Set Domain.'
    $RenameNameForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    $RenameNameFormBox1 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox1.Location = '10,10'
    $RenameNameFormBox1.size = '440,170'
    $RenameNameFormBox1.text = 'Current settings:'
    $RenameNameFormtextLabel1 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel1.Location = '20,20'
    $RenameNameFormtextLabel1.size = '100,20'
    $RenameNameFormtextLabel1.Text = 'UserName:' 
    $RenameNameFormtextLabel2 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel2.Location = '20,50'
    $RenameNameFormtextLabel2.size = '100,20'
    $RenameNameFormtextLabel2.Text = 'Display Name:'
    $RenameNameFormtextLabel3 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel3.Location = '20,80'
    $RenameNameFormtextLabel3.size = '100,20'
    $RenameNameFormtextLabel3.Text = 'First Name:'  
    $RenameNameFormtextLabel4 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel4.Location = '20,110'
    $RenameNameFormtextLabel4.size = '100,20'
    $RenameNameFormtextLabel4.Text = 'Last Name:' 
    $RenameNameFormtextLabel5 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel5.Location = '20,140'
    $RenameNameFormtextLabel5.size = '100,20'
    $RenameNameFormtextLabel5.Text = 'Email:'  
    $RenameNameFormtext1 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext1.Location = '120,20'
    $RenameNameFormtext1.size = '150,20'
    $RenameNameFormtext1.Text = $UserSamAccountName
    $RenameNameFormtext1.ForeColor = 'Blue'
    $RenameNameFormtext2 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext2.Location = '120,50'
    $RenameNameFormtext2.size = '200,20'
    $RenameNameFormtext2.Text = $SelectedUser
    $RenameNameFormtext2.ForeColor = 'Blue'
    $RenameNameFormtext3 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext3.Location = '120,80'
    $RenameNameFormtext3.size = '150,20'
    $RenameNameFormtext3.Text = $UserFirstName  
    $RenameNameFormtext3.ForeColor = 'Blue'
    $RenameNameFormtext4 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext4.Location = '120,110'
    $RenameNameFormtext4.size = '200,20'
    $RenameNameFormtext4.Text = $UserSurName 
    $RenameNameFormtext4.ForeColor = 'Blue'
    $RenameNameFormtext5 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext5.Location = '120,140'
    $RenameNameFormtext5.size = '300,20'
    $RenameNameFormtext5.Text = $UserEmail  
    $RenameNameFormtext5.ForeColor = 'Blue'
    $RenameNameFormBox2 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox2.Location = '460,10'
    $RenameNameFormBox2.size = '440,170'
    $RenameNameFormBox2.text = 'New Settings:'
    $RenameNameFormB2textLabel1 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel1.size = '100,20'
    $RenameNameFormB2textLabel1.Location = '20,20'
    $RenameNameFormB2textLabel1.Text = 'UserName:' 
    $RenameNameFormB2textLabel2 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel2.Location = '20,50'
    $RenameNameFormB2textLabel2.size = '100,20'
    $RenameNameFormB2textLabel2.Text = 'Display Name:'
    $RenameNameFormB2textLabel3 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel3.Location = '20,80'
    $RenameNameFormB2textLabel3.size = '100,20'
    $RenameNameFormB2textLabel3.Text = 'First Name:'  
    $RenameNameFormB2textLabel4 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel4.Location = '20,110'
    $RenameNameFormB2textLabel4.size = '100,20'
    $RenameNameFormB2textLabel4.Text = 'Last Name:' 
    $RenameNameFormB2text1 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text1.Location = '120,20'
    $RenameNameFormB2text1.size = '150,20'
    $RenameNameFormB2text1.ForeColor = 'Green'
    $RenameNameFormB2text1.Text = $UserSamAccountName
    $RenameNameFormB2text2 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text2.Location = '120,50'
    $RenameNameFormB2text2.size = '200,20'
    $RenameNameFormB2text2.Text = $SelectedUser
    $RenameNameFormB2text2.ForeColor = 'Green'
    $RenameNameFormB2text3 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text3.Location = '120,80'
    $RenameNameFormB2text3.size = '150,20'
    $RenameNameFormB2text3.Text = $UserFirstName  
    $RenameNameFormB2text3.ForeColor = 'Green'
    $RenameNameFormB2text4 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text4.Location = '120,110'
    $RenameNameFormB2text4.size = '200,20'
    $RenameNameFormB2text4.ForeColor = 'Green'
    $RenameNameFormB2text4.Text = $UserSurName 
    $RenameNameFormBox3 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox3.Location = '10,190'
    $RenameNameFormBox3.size = '890,100'
    $RenameNameFormBox3.text = 'Change Domain:'
    $BoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton1.Location = '20,30'
    $BoxRadioButton1.size = '200,35'
    $BoxRadioButton1.Checked = $true 
    $BoxRadioButton1.Text = 'SCTS.'
    $BoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton2.Location = '20,60'
    $BoxRadioButton2.size = '200,35'
    $BoxRadioButton2.Checked = $false
    $BoxRadioButton2.Text = 'Tribs.'
    $BoxRadioButton2.Add_Click( {
            $BoxRadioButton1.Checked = $false
            $BoxRadioButton2.Checked = $true })
    $BoxRadioButton1.Add_Click( {
            $BoxRadioButton1.Checked = $true
            $BoxRadioButton2.Checked = $false })
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '300,450'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '400,450'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            Write-Output "Email selection closed"
            $RenameNameForm.Close()
            $RenameNameForm.Dispose()
            Return MainForm })
    # Add all the Form controls on one line 
    $RenameNameForm.Controls.AddRange(@($RenameNameFormBox1, $RenameNameFormBox3, $RenameNameFormBox2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $RenameNameFormBox1.Controls.AddRange(@($RenameNameFormtextLabel1, $RenameNameFormtextLabel2, $RenameNameFormtextLabel3, $RenameNameFormtextLabel4, $RenameNameFormtextLabel5, $RenameNameFormtext1, $RenameNameFormtext2, $RenameNameFormtext3, $RenameNameFormtext4, $RenameNameFormtext5))
    $RenameNameFormBox2.Controls.AddRange(@($RenameNameFormB2textLabel1, $RenameNameFormB2textLabel2, $RenameNameFormB2textLabel3, $RenameNameFormB2textLabel4, $RenameNameFormB2textLabel5, $RenameNameFormB2text1, $RenameNameFormB2text2, $RenameNameFormB2text3, $RenameNameFormB2text4, $RenameNameFormB2text5))
    $RenameNameFormBox3.Controls.AddRange(@($BoxRadioButton1, $BoxRadioButton2))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $RenameNameForm.AcceptButton = $OKButton
    $RenameNameForm.CancelButton = $CancelButton
    # Activate the form
    $RenameNameForm.Add_Shown( { $RenameNameForm.Activate() })    
    # Get the results from the button click
    $Result = $RenameNameForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        if ($BoxRadioButton1.Checked) {
            $NewEmail1 = "$UserSamAccountName@ScotCourts.gov.uk"
            $NewEmail2 = "$UserSamAccountName@ScotCourtsTribunals.gov.uk"
            $NewEmail3 = "$UserSamAccountName@ScotCourts.pnn.gov.uk"
            $NewEmail4 = "$UserSamAccountName@ScotCourtsTribunals.pnn.gov.uk"
            $NewEmail5 = "$UserSamAccountName@scotcourtsgovuk.mail.onmicrosoft.com"
            Write-output "$UserSamAccountName@ScotCourts.gov.uk selected as primary"
            ######################################################################
            $newProxy = "smtp:" + $NewEmail2
            $NewPrimary = "SMTP:" + $NewEmail1
            $newProxy1 = "smtp:" + $NewEmail3
            $newProxy2 = "smtp:" + $NewEmail4
            $newProxy3 = "smtp:" + $NewEmail5
            Set-ADUser -identity $UserSamAccountName -EmailAddress "$UserSamAccountName@ScotCourts.gov.uk"
            Set-ADUser -identity $UserSamAccountName -replace @{proxyAddresses = ($NewPrimary) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy1) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy2) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy3) }
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show("The user $UserSamAccountName`nhas been Change to SCTS Domain.", 'Domain.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
        else {
            $NewEmail1 = "$UserSamAccountName@ScotCourtsTribunals.gov.uk"
            $NewEmail2 = "$UserSamAccountName@ScotCourts.gov.uk"
            $NewEmail3 = "$UserSamAccountName@ScotCourts.pnn.gov.uk"
            $NewEmail4 = "$UserSamAccountName@ScotCourtsTribunals.pnn.gov.uk"
            $NewEmail5 = "$UserSamAccountName@scotcourtsgovuk.mail.onmicrosoft.com"
            Write-output "$UserSamAccountName@ScotCourtsTribunals.gov.uk selected as primary"
            ######################################################################
            $newProxy = "smtp:" + $NewEmail2
            $NewPrimary = "SMTP:" + $NewEmail1
            $newProxy1 = "smtp:" + $NewEmail3
            $newProxy2 = "smtp:" + $NewEmail4
            $newProxy3 = "smtp:" + $NewEmail5
            Set-ADUser -identity $UserSamAccountName -EmailAddress "$UserSamAccountName@ScotCourtsTribunals.gov.uk"
            Set-ADUser -identity $UserSamAccountName -replace @{proxyAddresses = ($NewPrimary) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy1) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy2) }
            Set-ADUser -identity $UserSamAccountName -add @{proxyAddresses = ($newProxy3) }
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show("The user $UserSamAccountName`nhas been Change to Tribs Domain.", 'Domain.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
    }
}

function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $EmailManForm = New-Object System.Windows.Forms.Form
    $EmailManForm.width = 550
    $EmailManForm.height = 375
    $EmailManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $EmailManForm.MinimizeBox = $False
    $EmailManForm.MaximizeBox = $False
    $EmailManForm.FormBorderStyle = 'Fixed3D'
    $EmailManForm.Text = $WinTitle
    $EmailManForm.Icon = $Icon
    $EmailManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    $Logo = [System.Drawing.Image]::Fromfile('\\saufs01\IT\Enterprise Team\Usermanagement\icons\SCTS.png')
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width = $Logo.Size.Width
    $pictureBox.Height = $Logo.Size.Height
    $pictureBox.Image = $Logo
    $EmailManForm.controls.add($pictureBox)
    $EmailManFormtext1 = New-Object System.Windows.Forms.Label
    $EmailManFormtext1.Location = '20,120'
    $EmailManFormtext1.size = '500,50'
    $EmailManFormtext1.Text = "This script changes the following on a Users account: `n - AD account - Changes users Primary email address."
    $EmailBox1 = New-Object System.Windows.Forms.GroupBox
    $EmailBox1.Location = '10,175'
    $EmailBox1.size = '500,75'
    $EmailBox1.text = '1. Select a UserName from the dropdown lists:'
    $EmailManFormtextLabel1 = New-Object System.Windows.Forms.Label
    $EmailManFormtextLabel1.Location = '20,40'
    $EmailManFormtextLabel1.size = '100,20'
    $EmailManFormtextLabel1.Text = 'UserName:' 
    $EmailManFormNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $EmailManFormNameComboBox1.Location = '125,35'
    $EmailManFormNameComboBox1.Size = '350, 310'
    $EmailManFormNameComboBox1.AutoCompleteMode = 'Suggest'
    $EmailManFormNameComboBox1.AutoCompleteSource = 'ListItems'
    $EmailManFormNameComboBox1.Sorted = $true;
    $EmailManFormNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $EmailManFormNameComboBox1.DataSource = $UsernameList
    $EmailManFormNameComboBox1.add_SelectedIndexChanged( { $ChangeThisUser.Text = "$($EmailManFormNameComboBox1.SelectedItem.ToString())" })
    $EmailManFormtext2 = New-Object System.Windows.Forms.Label
    $EmailManFormtext2.Location = '20,275'
    $EmailManFormtext2.size = '75,150'
    $EmailManFormtext2.Text = 'Change:'
    $ChangeThisUser = New-Object System.Windows.Forms.Label
    $ChangeThisUser.Location = '100,275'
    $ChangeThisUser.Size = '200,50'
    $ChangeThisUser.ForeColor = 'Blue'
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '300,280'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '400,280'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
        Write-output "Email selection closed"
            $EmailManForm.Close()
            $EmailManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $EmailManForm.Controls.AddRange(@($EmailBox1, $ChangeThisUser, $EmailManFormtext1, $EmailManFormtext2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $EmailBox1.Controls.AddRange(@($EmailManFormtextLabel1, $EmailManFormNameComboBox1))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $EmailManForm.AcceptButton = $OKButton
    $EmailManForm.CancelButton = $CancelButton
    # Activate the form
    $EmailManForm.Add_Shown( { $EmailManForm.Activate() })    
    # Get the results from the button click
    $Result = $EmailManForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        if ($ChangeThisUser.Text -eq '') {
            Write-Output "no user selected"
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Renamer.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
        $SelectedUser = $ChangeThisUser.text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'SamAccountName'
        $UserFirstName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'GivenName'
        $UserSurName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'Surname'
        $UserEmail = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" -Properties * | Select-Object -ExpandProperty EmailAddress
        Write-output "$UserSamAccountName Selected"
        EmailForm
    }
}
MainForm
