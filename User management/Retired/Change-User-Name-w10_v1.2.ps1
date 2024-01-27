# User Management - Change of Name v1.1
# Author         John Mckay
# Date           17/08/2018
# Version        1.1 (2.02)
# Purpose        To change username & email address.
# Useage         Helpdesk staff run a shorcut to Change of Name.exe
#                Script loads exchange 2013 session & imports users from AD.Opens form & helpdesk select current username and enter
#                new surname and new logonname. Email is sent to IT Enterprise to delete old username email addresses in a month.
#
# Changes        v2.02 05/12/2019 GK changes to take account of the new personal folder location. 
#                                 Added 'ou=SCTS Users,ou=User Accounts,ou=SCTS,DC=scotcourts,DC=local' as another searchbase.
#                v2.01 17/08/18 - changed emailadderss for message2 & message3 to OldUserNameSam instead of newsamname.
#                v2.00 16/08/2018 updated form to use popup messages to give helpdesk status messages & uses try & catch. 
#                v1.01 14/02/2018 JM changed send email to helpdesk to use $env:UserName instead of helpdesk.
#
# Script function: This script performs the following on a User account.
$message = "
This script changes the following on a Users account:

AD account - Changes users logon name.
AD account - Changes users displayname.
AD account - Changes users profile path.
AD account - Changes users P drive path.
SAUFS01 - Renames users P drive folder name.
SAUFS01 - Renames users profile folder name.
Email - Changes users email address.

Please note: The user needs to be logged off before running !!!"

$Version='1.2'
##  Create Session with Exchange 2013   ###
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session

##  Set icon for form   ###
$Icon = "\\scotcourts.local\data\IT\Enterprise Team\Usermanagement\icons\user.ico"

##  Get listof UserNames from AD OU's   ###
#region changeBlock
#$Users1 = Get-aduser –filter * -searchbase 'ou=tribunalusers,ou=tribunals,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users2 = Get-aduser –filter * -searchbase 'ou=sheriffsparttime,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users3 = Get-aduser –filter * -searchbase 'ou=scs employees,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Where-Object {($_.DistinguishedName -notlike '*OU=deleted users,*') -and ($_.DistinguishedName -notlike '*OU=it administrators,*')} | Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users4 = Get-aduser –filter * -searchbase 'ou=JP,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users5 = Get-aduser –filter * -searchbase 'ou=judges,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users6 = Get-aduser –filter * -searchbase 'ou=sheriffs,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname| select-Object -ExpandProperty DisplayName
#$users7 = Get-aduser –filter * -searchbase 'ou=sheriffsprincipal,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users8 = Get-aduser –filter * -searchbase 'ou=sheriffssummary,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users9 = Get-aduser –filter * -searchbase 'ou=sheriffsretired,ou=scs users,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users10 = Get-aduser –filter * -searchbase 'ou=courts,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users11 = Get-aduser –filter * -searchbase 'ou=judiciary,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users12 = Get-aduser –filter * -searchbase 'ou=tribunals,ou=scts users,ou=useraccounts,ou=courts,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$users13 = Get-aduser –filter * -searchbase 'ou=SCTS Users,ou=User Accounts,ou=SCTS,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
$users1 = Get-aduser –filter * -searchbase 'ou=SCTS Users,ou=User Accounts,ou=SCTS,DC=scotcourts,DC=local' -Properties DisplayName| Select-Object Displayname | select-Object -ExpandProperty DisplayName
#$UserNameList = $Users1 + $users2 + $users3 + $users4 + $users5 + $users6 + $users7 + $users8 + $users9 + $users10 + $users11 + $users12 + $users13
$UserNameList = $Users1 
#endregion changeBlock

##  Show Start Message:   ###
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", 'Change Of User Name v2.00', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {Exit} 
else {
    ## create the pop up information form   ###
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        # create form
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = "Change Of User Name v2.00"
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
    
    ##  Create ChangeOfName Form   ###
    Function ChangeOfName {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        ##  Define form details   ###
        $ChangeNameForm = New-Object System.Windows.Forms.Form
        $ChangeNameForm.width = 780
        $ChangeNameForm.height = 500
        $ChangeNameForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $ChangeNameForm.MinimizeBox = $False
        $ChangeNameForm.MaximizeBox = $False
        $ChangeNameForm.Icon = $Icon
        $ChangeNameForm.FormBorderStyle = 'Fixed3D'
        $ChangeNameForm.Controlbox = $false
        $ChangeNameForm.Text = "User Management - Change of Name (W10) v$Version"
        $ChangeNameForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $ChangeNameForm.Topmost = $True

        ##  Add Ok button   ###               
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(610, 375)
        $OKButton.Size = New-Object System.Drawing.Size(100, 40)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ChangeNameForm.AcceptButton = $OKButton

        ##  Add Cancel button   ###   
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(490, 375)
        $CancelButton.Size = New-Object System.Drawing.Size(100, 40)
        $CancelButton.Text = 'Exit'
        $CancelButton.Add_Click( {
                $ChangeNameForm.Close
                $ChangeNameForm.Dispose
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        $ChangeNameForm.CancelButton = $CancelButton

        ##  Create group 1 box   ###
        $ChangeNameBox1 = New-Object System.Windows.Forms.GroupBox
        $ChangeNameBox1.Location = '40,30'
        $ChangeNameBox1.size = '700,80'
        $ChangeNameBox1.text = '1. Select a UserName from the dropdown list:'

        ##  Create group 1 box text label   ###
        $UserNameLabel1 = New-Object System.Windows.Forms.Label
        $UserNameLabel1.Location = '20,40'
        $UserNameLabel1.size = '150,30'
        $UserNameLabel1.Text = 'UserName:' 

        ##  Create group 1 box combo boxes   ###
        $UserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
        $UserNameComboBox1.Location = '420,35'
        $UserNameComboBox1.Size = '250, 40'
        $UserNameComboBox1.AutoCompleteMode = 'Suggest'
        $UserNameComboBox1.AutoCompleteSource = 'ListItems'
        $UserNameComboBox1.Sorted = $true;
        $UserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $UserNameComboBox1.DataSource = $UserNameList
        $UserNameComboBox1.add_SelectedIndexChanged( {$UserNameCombo1OutLabel4.Text = "$($UserNameComboBox1.SelectedItem.ToString())"})

        ##  Create label for combobox1 output   ### 
        $UserNameCombo1OutLabel4 = New-Object System.Windows.Forms.Label
        $UserNameCombo1OutLabel4.Location = '50,800'
        $UserNameCombo1OutLabel4.size = '250,30'
        $ChangeNameBox1.Controls.AddRange(@($UserNameLabel1, $UserNameComboBox1, $UserNameCombo1OutLabel4))

        ##  Create group 2 box in form   ###
        $ChangeNameBox2 = New-Object System.Windows.Forms.GroupBox
        $ChangeNameBox2.Location = '40,140'
        $ChangeNameBox2.size = '700,125'
        $ChangeNameBox2.text = '2. Enter new Surname and new LogOnName:'

        ##  Create group 2 box text labels   ###
        $SurnameLabel2 = New-Object System.Windows.Forms.Label
        $SurnameLabel2.Location = '20,40'
        $SurnameLabel2.size = '300,40'
        $SurnameLabel2.Text = 'Enter New Surname: (e.g Bloggs):' 
        $LogOnNameLabel3 = New-Object System.Windows.Forms.Label
        $LogOnNameLabel3.Location = '20,80'
        $LogOnNameLabel3.size = '300,40'
        $LogOnNameLabel3.Text = 'Enter New LogOnName: (e.g jbloggs):'

        ##  Create group 2 box text boxes   ###        
        $SurnametextBox1 = New-Object System.Windows.Forms.TextBox
        $SurnametextBox1.Location = '420,35'
        $SurnametextBox1.Size = '250,40'
        $LogOnNametextBox2 = New-Object System.Windows.Forms.TextBox
        $LogOnNametextBox2.Location = '420,75'
        $LogOnNametextBox2.Size = '250,40'
        $ChangeNameBox2.Controls.AddRange(@($SurnameLabel2, $LogOnNameLabel3, $SurnametextBox1, $LogOnNametextBox2))

        ##  Add all the Form controls   ### 
        $ChangeNameForm.Controls.AddRange(@($ChangeNameBox1, $ChangeNameBox2, $OKButton, $CancelButton))
        $ChangeNameForm.AcceptButton = $OKButton
        $ChangeNameForm.CancelButton = $CancelButton

        ##### Add all the GroupBox controls ###
        $ChangeNameBox1.Controls.AddRange(($UserNameLabel1, $UserNameComboBox1))
        $ChangeNameBox2.Controls.AddRange(($SurnameLabel2, $LogOnNameLabel3, $SurnametextBox1, $LogOnNametextBox2))

        ## Completed - Create ChangeOfName Form   ###
        ##  set return values   ###
        $result = $ChangeNameForm.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {

            ##  Check textboxes have values   ###
            $NewSurname = $SurnametextBox1.Text
            $NewLogOnName = $LogOnNametextBox2.Text
            $CurrentUserName = $UserNameCombo1OutLabel4.Text 

            ##  CHECK - Don't accept no User selection   ### 
            if ($UserNameCombo1OutLabel4.Text -eq "") {
                [System.Windows.Forms.MessageBox]::Show("You need to select a User !!!!!`n`nTrying to enter blank fields is never a good idea.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                break
            }
            if ($NewSurname -eq "") {
                [System.Windows.Forms.MessageBox]::Show("You need to enter a new Surname!", "User Management - Change of Name.")
                Return ChangeOfName
            }
            if ($NewLogOnName -eq "") {
                [System.Windows.Forms.MessageBox]::Show("You need to enter a new LogOnName!", "User Management - Change of Name.")
                Return ChangeOfName
            }
            ##  Check to see if New LogOnName is already in use   ###
            if ($NewLogOnName -ne "") {
                $User = Get-ADUser -Filter {sAMAccountName -eq $NewLogOnName}
                If ($User -ne $Null) {
                    Add-Type -AssemblyName System.Windows.Forms 
                    [System.Windows.Forms.MessageBox]::Show("The LogOnName - $NewLogOnName - can't be used because it's assigned to an existing user account.`n`nThe next page will display the current usernames in use.`n`nPlease use a LogOnName that's not currently in use.", "ERROR - CAN'T USE NEW NAME", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    Get-AdUser -Filter "SamAccountName -like '$NewLogOnName*'" | Where-Object {($_.SamAccountName -notlike '*admin')} | Select-Object SamAccountName | Out-GridView -Title "List of existing user names currently in use"
                    $ChangeNameForm.Close()
                    $ChangeNameForm.Dispose()
                    ChangeOfName
                }
            }
            
            ## set variables for first name & sam name   ###
            $OldUserName = Get-ADUser -Filter "Name -eq '$CurrentUserName'" -properties *
            $FirstName = $OldUserName.GivenName
            $OldUserNameSam = $OldUserName.SamAccountname
            
            ##  check only 1 samaccountname   ###
            if (($OldUserNameSam).Count -ne 1) {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG !!!.`n`nThere seems to be more than 1 user with the DisplayName.`n`nCheck the Users in AD & change the DisplayNames so they are unique.`n`ne.g. change - Bloggs, Joe - to - Bloggs, Joe (Glasgow) -.`n`nIf you need asisstance to do this please contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Return ChangeOfName  
            }
            
            ## change users profile folder name  ###
            $poplabel = "Renaming the users Profile folder`n`non \\scotcourts.local\data\profiles."
            PopupForm
            try {
                if (test-path \\scotcourts.local\data\profiles\"$OldUserNameSam.v2") {
                    Rename-Item \\scotcourts.local\data\profiles\"$OldUserNameSam.v2" -NewName "$NewLogOnName.v2" -ErrorAction Stop
                    Write-Verbose "Renamed users Profile folder" -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG with renaming the users Profile folder !!!.`n`nCheck the User isn't logged on !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
                   
            ## change users P drive folder name  ###
            $poplabel = "Renaming the users P drive folder`n`non \\scotcourts.local\home\P."
            PopupForm
            try {
                if (test-path \\scotcourts.local\home\P\$OldUserNameSam) {
                    Rename-Item \\scotcourts.local\home\P\$OldUserNameSam -NewName "$NewLogOnName" -ErrorAction Stop
                    Write-Verbose "Renamed Users P drive folder" -Verbose
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG with renaming the users P drive folder !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
            
            ##  Change users details in AD   ###
            $poplabel = "Renaming the following users AD properties:`n`nDisplayname.`nLogOn name.`nProfile path.`nP drive path."
            PopupForm
            try {
                #region    RemoveHomepathFromSet-ADUserParameters                
                #Set-AdUser -Passthru -Identity $OldUserNameSam -SamAccountName $NewLogOnName -UserPrincipalName $NewLogOnName@scotcourts.local -Surname $NewSurname -HomeDirectory \\saufs01\users\$NewLogOnName -ProfilePath \\saufs01\profiles\$NewLogOnName -DisplayName ("$NewSurname, $FirstName")
                if ($null -eq $OldUserName.ProfilePath) {
                    Set-AdUser -Passthru -Identity $OldUserNameSam -SamAccountName $NewLogOnName -UserPrincipalName $NewLogOnName@scotcourts.local -Surname $NewSurname -DisplayName ("$NewSurname, $FirstName")
                }
                else {
                    Set-AdUser -Passthru -Identity $OldUserNameSam -SamAccountName $NewLogOnName -UserPrincipalName $NewLogOnName@scotcourts.local -Surname $NewSurname -ProfilePath \\scotcourts.local\data\profiles\$NewLogOnName -DisplayName ("$NewSurname, $FirstName")
                }
                #endregion RemoveHomepathFromSet-ADUserParameters 
                Write-Verbose "Changed Users AD properties" -Verbose
            }
            catch {                
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG changing the Users AD properties !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
            
            ##  Change users mailbox name & email address   ###
            $poplabel = "Renaming the users mailbox and email address."
            PopupForm
            try {
                Set-Mailbox -Identity $OldUserNameSam -Alias $NewLogOnName -Name ("$NewSurname, $FirstName")
                Write-Verbose "Changed Users mailbox name and email address." -Verbose
            }
            catch {                
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG changing the Users mailbox name & email address !!!.`n`nPlease contact the Systems Integration Team with the details.", 'User Management - Change of User Name.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $ChangeNameForm.Close()
                $ChangeNameForm.Dispose()
                Return ChangeOfName
            }
            
            ##  Send email1 to user re change of name  ###
            $message2 = "
                The user $CurrentUserName has had a change of name to $NewSurname, $FirstName.`n
                The old $CurrentUserName email address will be valid for 1 month.`n
                During this time any emails sent to $CurrentUserName will be delivered into your mailbox.`n
                Regards IT helpdesk."
            Send-MailMessage -To $OldUserNameSam@scotcourts.gov.uk -from helpdesk@scotcourts.gov.uk -Subject "Change of Name - The user $CurrentUserName email address." -Body "$message2" -SmtpServer mail.scotcourts.local
            
            ##  Send email2 to user re outlook reset  ###
            $message3 = "
                Your logon name and email address have been changed to your new name.`n
                Outlook needs to be reset on your computer to pick up your new name.`n
                You can do this by following the attached instructions`n
                or`n
                contact the IT helpdesk with your computer tag number requesting your Outlook profile to be reset.`n
                PLEASE NOTE: When you reset your outlook profile any shared mailboxes you have access to will take a minute or 2 to reconnect.`n
                Regards IT helpdesk."
            Send-MailMessage -To $OldUserNameSam@scotcourts.gov.uk -from helpdesk@scotcourts.gov.uk -Subject "Change of Name - Outlook needs manual reset to update to your new name." -Attachments "\\scotcourts.local\data\it\Enterprise Team\UserManagement\User Info documents\reset outlook profile.docx" -Body "$message3" -SmtpServer mail.scotcourts.local
            
            ## Send confirmation email to helpdesk   ###
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: Change of Name - The user $OldUserNameSam has had a change of name to $NewLogOnName." -Body "$message2" -SmtpServer mail.scotcourts.local
            
            ##  Send email to enterprise to remove old email addresses   ###
            $message4 = "
                The user $OldUserNameSam has had a change of name to $NewLogOnName.`n
                The old $OldUserNameSam email addresses need removed from their account."
            $Date = (Get-Date).AddDays(31).ToString('dd-MMMM') 
            Send-MailMessage -To itenterprise@scotcourts.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "Change of Name - $OldUserNameSam - Remove old email addresses $Date" -Body "$message4" -SmtpServer mail.scotcourts.local
            
            ## complete message   ###
            $message5 = "
                `nThe User $OldUserNameSam has had a Change of Name to $NewLogOnName.
                `nThe user can now logon as $NewLogOnName with their existing password.
                `nThe User needs to reset their Outlook profile on their computer to pick up the new name and an email has been sent to the User with details of how to do this.
                `nAn email has been sent to IT Infrastructure to remove the $OldUserNameSam email addresses in 1 month."
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("$message5", "Change of Name - Completed.", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) 
            ChangeOfName  
            
        }   
    }
}
ChangeOfName
