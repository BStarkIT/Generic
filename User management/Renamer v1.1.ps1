# Rename User Account V1.1
# Author        Brian Stark
# Date          04/06/2021
# Version       1.1
# Purpose       To Automate the rename procudure.
# Useage        helpdesk staff run a shorcut to Users - Archive Leaver Accounts.
#               Fork of the disable 
#
# Changes       V1.1 04/06/2021 BS Clean up & commented
#               V1.0 28/12/2020 first release
#               
#
# Script function: This script performs the following on a User account:
#               Rename Sam
#               Rename P folder
#               Replace email addresses (verification required)
#               Replace Display name
#               Replace Surname
#
#   Script does NOT:
#               replace Wifi Cert
#               update digital signature in Adobe
#
# Set icon for all forms and subforms
#
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\User.ico'
#
# Get listof UserNames from AD OU's
#
$Users1 = Get-ADUser –filter * -searchbase 'ou=tribunalusers,ou=tribunals,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users2 = Get-ADUser –filter * -searchbase 'ou=sheriffsparttime,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users3 = Get-ADUser –filter * -searchbase 'ou=scs employees,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Where-Object { ($_.DistinguishedName -notlike '*OU=deleted users,*') -and ($_.DistinguishedName -notlike '*OU=it administrators,*') } | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users4 = Get-ADUser –filter * -searchbase 'ou=JP,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users5 = Get-ADUser –filter * -searchbase 'ou=judges,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users6 = Get-ADUser –filter * -searchbase 'ou=sheriffs,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users7 = Get-ADUser –filter * -searchbase 'ou=sheriffsprincipal,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users8 = Get-ADUser –filter * -searchbase 'ou=sheriffssummary,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users9 = Get-ADUser –filter * -searchbase 'ou=sheriffsretired,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users10 = Get-ADUser –filter * -searchbase 'ou=courts,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users11 = Get-ADUser –filter * -searchbase 'ou=judiciary,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users12 = Get-ADUser –filter * -searchbase 'ou=tribunals,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$users13 = Get-ADUser –filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$UserNameList = $Users1 + $users2 + $users3 + $users4 + $users5 + $users6 + $users7 + $users8 + $users9 + $users10 + $users11 + $users12 + $users13
#
function RenameForm {
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
    $RenameNameForm.Text = 'Rename User account.'
    $RenameNameForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    #
    $RenameNameFormBox1 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox1.Location = '10,10'
    $RenameNameFormBox1.size = '440,200'
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
    $RenameNameFormtextLabel5.Text = 'P Drive:'
    $RenameNameFormtextLabel6 = New-Object System.Windows.Forms.Label
    $RenameNameFormtextLabel6.Location = '20,170'
    $RenameNameFormtextLabel6.size = '100,20'
    $RenameNameFormtextLabel6.Text = 'Email:'  
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
    $RenameNameFormtext5.size = '250,20'
    $RenameNameFormtext5.Text = "\\SauFS01\p\$UserSamAccountName"
    $RenameNameFormtext5.ForeColor = 'Blue'
    $RenameNameFormtext6 = New-Object System.Windows.Forms.Label
    $RenameNameFormtext6.Location = '120,170'
    $RenameNameFormtext6.size = '300,20'
    $RenameNameFormtext6.Text = $UserEmail  
    $RenameNameFormtext6.ForeColor = 'Blue'
    #
    $RenameNameFormBox2 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox2.Location = '460,10'
    $RenameNameFormBox2.size = '440,200'
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
    $RenameNameFormB2textLabel5 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel5.Location = '20,140'
    $RenameNameFormB2textLabel5.size = '100,20'
    $RenameNameFormB2textLabel5.Text = 'P Drive:'
    $RenameNameFormB2textLabel6 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2textLabel6.Location = '20,170'
    $RenameNameFormB2textLabel6.size = '100,20'
    $RenameNameFormB2textLabel6.Text = 'Email:'  
    $RenameNameFormB2text1 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text1.Location = '120,20'
    $RenameNameFormB2text1.size = '150,20'
    $RenameNameFormB2text1.ForeColor = 'Green'
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
    $RenameNameFormB2text5 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text5.Location = '120,140'
    $RenameNameFormB2text5.size = '250,20'
    $RenameNameFormB2text5.ForeColor = 'Green'
    $RenameNameFormB2text6 = New-Object System.Windows.Forms.Label
    $RenameNameFormB2text6.Location = '120,170'
    $RenameNameFormB2text6.size = '300,20'
    $RenameNameFormB2text6.ForeColor = 'Green'
    #
    # Text boxes
    $RenameNameFormBox3 = New-Object System.Windows.Forms.GroupBox
    $RenameNameFormBox3.Location = '10,220'
    $RenameNameFormBox3.size = '890,100'
    $RenameNameFormBox3.text = 'Change:'
    #
    $SurNameBoxText = New-Object System.Windows.Forms.Label
    $SurNameBoxText.Location = '20,30'
    $SurNameBoxText.size = '75,20'
    $SurNameBoxText.Text = "Last Name" 
    $UserNameBoxText = New-Object System.Windows.Forms.Label
    $UserNameBoxText.Location = '20,60'
    $UserNameBoxText.size = '75,20'
    $UserNameBoxText.Text = "UserName"  
    $SurNameBox = New-Object System.Windows.Forms.TextBox
    $SurNameBox.Location = New-Object System.Drawing.Point(100, 30)
    $SurNameBox.Size = New-Object System.Drawing.Size(330, 20)
    $SurNameBox.add_TextChanged( { $RenameNameFormB2text4.Text = "$($SurNameBox.text)" })
    $UserNameBox = New-Object System.Windows.Forms.TextBox
    $UserNameBox.Location = New-Object System.Drawing.Point(100, 60)
    $UserNameBox.Size = New-Object System.Drawing.Size(330, 20)
    $UserNameBox.add_TextChanged( { $RenameNameFormB2text1.Text = $UserNameBox.text })
    $UserNameBox.add_TextChanged( { $RenameNameFormB2text5.Text = "\\SauFS01\p\" + $UserNameBox.text })
    $UserNameBox.add_TextChanged( { $RenameNameFormB2text6.Text = $UserNameBox.text + "@Scotcourts.gov.uk" })
    $SurNameBox.add_TextChanged( { $RenameNameFormB2text2.Text = $SurNameBox.text + ", " + $UserFirstName })
    #
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
            $RenameNameForm.Close()
            $RenameNameForm.Dispose()
            Return MainForm })
    # Add all the Form controls on one line 
    $RenameNameForm.Controls.AddRange(@($RenameNameFormBox1, $RenameNameFormBox3, $RenameNameFormBox2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $RenameNameFormBox1.Controls.AddRange(@($RenameNameFormtextLabel1, $RenameNameFormtextLabel2, $RenameNameFormtextLabel3, $RenameNameFormtextLabel4, $RenameNameFormtextLabel5, $RenameNameFormtextLabel6, $RenameNameFormtext1, $RenameNameFormtext2, $RenameNameFormtext3, $RenameNameFormtext4, $RenameNameFormtext5, $RenameNameFormtext6))
    $RenameNameFormBox2.Controls.AddRange(@($RenameNameFormB2textLabel1, $RenameNameFormB2textLabel2, $RenameNameFormB2textLabel3, $RenameNameFormB2textLabel4, $RenameNameFormB2textLabel5, $RenameNameFormB2textLabel6, $RenameNameFormB2text1, $RenameNameFormB2text2, $RenameNameFormB2text3, $RenameNameFormB2text4, $RenameNameFormB2text5, $RenameNameFormB2text6))
    $RenameNameFormBox3.Controls.AddRange(@($SurNameBox, $UserNameBox, $SurNameBoxText, $UserNameBoxText))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $RenameNameForm.AcceptButton = $OKButton
    $RenameNameForm.CancelButton = $CancelButton
    # Activate the form
    $RenameNameForm.Add_Shown( { $RenameNameForm.Activate() })    
    # Get the results from the button click
    $Result = $RenameNameForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        $NewSurname = $SurNameBox.text
        $NewSAM = $UserNameBox.text
        if ($NewSurname -eq "") {
            [System.Windows.Forms.MessageBox]::Show("You need to enter a new Surname!", "User Management - Change of Name.")
            Return RenameNameForm
        }
        if ($NewSAM -eq "") {
            [System.Windows.Forms.MessageBox]::Show("You need to enter a new LogOnName!", "User Management - Change of Name.")
            Return RenameNameForm
        }
        else {
            $tentativeSAM = $NewSAM
            if (Get-ADUser -Filter { SamAccountName -eq $NewSAM }) {    
                do {
                    $inc ++
                    $NewSAM = $tentativeSAM + [string]$inc
                } 
                until (-not (Get-ADUser -Filter { SamAccountName -eq $NewSAM }))
            }
            if ($UserEmail -like '*@ScotCourtsTribunals.gov.uk*') { 
                $NewEmail1 = "$NewSAM@ScotCourtsTribunals.gov.uk"
                $NewEmail2 = "$NewSAM@ScotCourts.gov.uk"
                $NewEmail3 = "$NewSAM@ScotCourts.pnn.gov.uk"
                $NewEmail4 = "$NewSAM@ScotCourtsTribunals.pnn.gov.uk"
            }
            elseif ($UserEmail -like '*@ScotCourts.gov.uk*') { 
                $NewEmail1 = "$NewSAM@ScotCourts.gov.uk"
                $NewEmail2 = "$NewSAM@ScotCourtsTribunals.gov.uk"
                $NewEmail3 = "$NewSAM@ScotCourts.pnn.gov.uk"
                $NewEmail4 = "$NewSAM@ScotCourtsTribunals.pnn.gov.uk"
            }
            ######################################################################
            $poplabel = "Renaming the users P drive folder`n`non \\scotcourts.local\home\P."
            PopupForm
            try {
                if (Test-Path \\scotcourts.local\home\P\$UserSamAccountName) {
                    Rename-Item \\scotcourts.local\home\P\$UserSamAccountName -NewName "$NewSAM" -ErrorAction Stop
                    Write-Verbose "Renamed Users P drive folder" -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG with renaming the users P drive folder !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return MainForm
            }
            $poplabel = "Renaming the user."
            PopupForm
            Write-Output $ID
            Write-Output $NewSAM
            try {
                $DisplayName = $NewSurname + ", " + $UserFirstName
                Get-ADUser -identity $ID  | Set-AdUser -Replace @{SamAccountName = $NewSAM } -ErrorAction Stop
                Get-ADUser -identity $ID  | Set-AdUser -Replace @{Surname = $NewSurname } -ErrorAction Stop
                Get-ADUser -identity $ID  | Set-AdUser -Replace @{DisplayName = $DisplayName } -ErrorAction Stop
            }
            catch {                
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG changing the Users AD.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
            #
            ##  Change users mailbox name & email address   ###
            $poplabel = "Renaming the users mailbox and email address."
            PopupForm

            try {
                Set-Mailbox -Identity $UserSamAccountName -Alias $NewSAM -Name ("$NewSurname, $UserFirstName")
                Set-ADUser -Identity $UserSamAccountName -emailaddress $NewEmail1
                Write-Verbose "Changed Users mailbox name and email address." -Verbose
            }
            catch {                
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG changing the Users mailbox name & email address !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
            $poplabel = "Adding new Aliases to the users mailbox."
            PopupForm
            $newProxy1 = "smtp:" + $NewEmail2
            $newProxy2 = "smtp:" + $NewEmail3
            $newProxy3 = "smtp:" + $NewEmail4
            $NewPrimary = "SMTP:" + $NewEmail1
            Set-ADUser -identity $ID -Add @{proxyAddresses = ($newProxy1) }
            Set-ADUser -identity $ID -Add @{proxyAddresses = ($newProxy2) }
            Set-ADUser -identity $ID -Add @{proxyAddresses = ($newProxy3) }
            Set-ADUser -identity $ID -Add @{proxyAddresses = ($NewPrimary) }
        }
        Return MainForm
    }
}
Function PopUpForm {
    Add-Type -AssemblyName System.Windows.Forms    
    # create form
    $PopForm = New-Object System.Windows.Forms.Form
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

Function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $RenameManForm = New-Object System.Windows.Forms.Form
    $RenameManForm.width = 550
    $RenameManForm.height = 500
    $RenameManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $RenameManForm.MinimizeBox = $False
    $RenameManForm.MaximizeBox = $False
    $RenameManForm.FormBorderStyle = 'Fixed3D'
    $RenameManForm.Text = 'User Renamer V3.0'
    $RenameManForm.Icon = $Icon
    $RenameManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    $Logo = [System.Drawing.Image]::Fromfile('\\saufs01\IT\Enterprise Team\Usermanagement\icons\SCTS.png')
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width = $Logo.Size.Width
    $pictureBox.Height = $Logo.Size.Height
    $pictureBox.Image = $Logo
    $RenameManForm.controls.add($pictureBox)
    $RenameManFormtext1 = New-Object System.Windows.Forms.Label
    $RenameManFormtext1.Location = '20,120'
    $RenameManFormtext1.size = '500,150'
    $RenameManFormtext1.Text = "This script changes the following on a Users account: `n - AD account - Changes users logon name.`n - AD account - Changes users displayname.`n - AD account - Changes users P drive path.`n - SAUFS01 - Renames users P drive folder name.`n - Email - Changes users email address.`nPlease note: The user needs to be logged off before running !!!"
    $RenameBox1 = New-Object System.Windows.Forms.GroupBox
    $RenameBox1.Location = '10,300'
    $RenameBox1.size = '500,75'
    $RenameBox1.text = '1. Select a UserName from the dropdown lists:'
    $RenameManFormtextLabel1 = New-Object System.Windows.Forms.Label
    $RenameManFormtextLabel1.Location = '20,40'
    $RenameManFormtextLabel1.size = '100,20'
    $RenameManFormtextLabel1.Text = 'UserName:' 
    $RenameManFormNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $RenameManFormNameComboBox1.Location = '125,35'
    $RenameManFormNameComboBox1.Size = '350, 310'
    $RenameManFormNameComboBox1.AutoCompleteMode = 'Suggest'
    $RenameManFormNameComboBox1.AutoCompleteSource = 'ListItems'
    $RenameManFormNameComboBox1.Sorted = $true;
    $RenameManFormNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $RenameManFormNameComboBox1.DataSource = $UsernameList
    $RenameManFormNameComboBox1.add_SelectedIndexChanged( { $RenameThisUser.Text = "$($RenameManFormNameComboBox1.SelectedItem.ToString())" })
    $RenameManFormtext2 = New-Object System.Windows.Forms.Label
    $RenameManFormtext2.Location = '20,420'
    $RenameManFormtext2.size = '75,150'
    $RenameManFormtext2.Text = 'Rename:'
    $RenameThisUser = New-Object System.Windows.Forms.Label
    $RenameThisUser.Location = '100,420'
    $RenameThisUser.Size = '200,50'
    $RenameThisUser.ForeColor = 'Blue'
    ### Add an OK button ###
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '300,400'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '400,400'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $RenameManForm.Close()
            $RenameManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $RenameManForm.Controls.AddRange(@($RenameBox1, $RenameThisUser, $RenameManFormtext1, $RenameManFormtext2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $RenameBox1.Controls.AddRange(@($RenameManFormtextLabel1, $RenameManFormNameComboBox1))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $RenameManForm.AcceptButton = $OKButton
    $RenameManForm.CancelButton = $CancelButton
    # Activate the form
    $RenameManForm.Add_Shown( { $RenameManForm.Activate() })    
    # Get the results from the button click
    $Result = $RenameManForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        if ($RenameThisUser.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Renamer.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
        $SelectedUser = $RenameThisUser.text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'SamAccountName'
        $UserFirstName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'GivenName'
        $UserSurName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'Surname'
        $UserEmail = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" -Properties * | Select-Object -ExpandProperty EmailAddress
        $ID = $((Get-ADUser $UserSamAccountName -Properties objectGUID).ObjectGUID).guid
        RenameForm
    }
}
MainForm

