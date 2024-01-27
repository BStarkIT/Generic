# M Mckenna Mailbox Management v1.00
# Author        John Mckay
# Date          19/07/2018
# Version       1.00
# Purpose       To manage access permissions on Martin Mckenna's mailbox.
# Useage        helpdesk staff run a shorcut & script loads exchange 2013 session & imports users from list.
#
# Changes 
#
$global:ErrorActionPreference = "Stop"
#####################################################################
#####    Create Session with Exchange 2013 & import user list   #####
#####################################################################
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$UsernameList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\lists\MMckenna.csv"
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\email.ico'
#################################################################
###############          Create Form1       #####################
#################################################################
Function Form1 {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ## Set the details of the form. ###
    $Form = New-Object System.Windows.Forms.Form
    $Form.size = '780,500'
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Form.Controlbox = $false
    $Form.Icon = $Icon
    $Form.FormBorderStyle = 'Fixed3D'
    $Form.Text = "$FormName"
    $Form.Font = New-Object System.Drawing.Font('Ariel', 10)
    ## Create group 1 box in form. ###
    $Box1 = New-Object System.Windows.Forms.GroupBox
    $Box1.Location = '40,40'
    $Box1.size = '700,125'
    $Box1.text = '1. Select a UserName from the dropdown list:'
    ## Create group 1 box text labels. ###
    $textLabel1 = New-Object System.Windows.Forms.Label;
    $textLabel1.Location = '20,40'
    $textLabel1.size = '150,40'
    $textLabel1.Text = 'UserName:' 
    ## Create group 1 box combo box. ###
    $UserNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $UserNameComboBox1.Location = '325,35'
    $UserNameComboBox1.Size = '350, 310'
    $UserNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $UserNameComboBox1.DataSource = $UsernameList.Name
    $UserNameComboBox1.add_SelectedIndexChanged( {$SelectedUserNametextLabel4.Text = "$($UserNameComboBox1.SelectedItem.ToString())"})
    ## Create group 2 box in form. ###
    $Box2 = New-Object System.Windows.Forms.GroupBox
    $Box2.Location = '40,190'
    $Box2.size = '700,125'
    $Box2.text = '2. Check the details below are correct before proceeding:'
    ## Create group 2 box text labels.  ###
    $textLabel3 = New-Object 'System.Windows.Forms.Label';
    $textLabel3.Location = '40,40'
    $textLabel3.size = '100,40'
    $textLabel3.Text = 'The User:' 
    $SelectedUserNametextLabel4 = New-Object System.Windows.Forms.Label
    $SelectedUserNametextLabel4.Location = '30,80'
    $SelectedUserNametextLabel4.Size = '200,40'
    $SelectedUserNametextLabel4.ForeColor = 'Blue'
    $textLabel5 = New-Object 'System.Windows.Forms.Label';
    $textLabel5.Location = '275,40'
    $textLabel5.size = '700,40'
    #$textLabel5.size = '400,40'
    $textLabel5.Text = "$Text5"
    ## Create group 3 box in form. ###
    $Box3 = New-Object System.Windows.Forms.GroupBox
    $Box3.Location = '40,340'
    $Box3.size = '700,30'
    $Box3.text = "$TextBox3"
    $Box3.button
    ## Add an OK button ###
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '640,390'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ## Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,390'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to Form'
    $CancelButton.add_Click( {
            $Form.Close()
            $Form.Dispose()
            Return MainForm})
    ## Add all the Form controls ### 
    $Form.Controls.AddRange(@($Box1, $Box2, $Box3, $OKButton, $CancelButton))
    ## Add all the GroupBox controls ###
    $Box1.Controls.AddRange(@($textLabel1, $UserNameComboBox1))
    $Box2.Controls.AddRange(@($textLabel3, $SelectedUserNametextLabel4, $textLabel5, $SelectedMailBoxNametextLabel6))
    ## Assign the Accept and Cancel options in the form ### 
    $Form.AcceptButton = $OKButton
    $Form.CancelButton = $CancelButton
    ## Activate the form ###
    $Form.Add_Shown( {$Form.Activate()})    
    ## Get the results from the button click ###
    $dialogResult = $Form.ShowDialog()
    ## If the OK button is selected
    if ($dialogResult -eq 'OK') {
        ##  Don't accept null username box   ### 
        if ($SelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!`n`nTrying to enter blank fields is never a good idea.", "$FormName", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $Form.Close()
            $Form.Dispose()
            break
        }
        ##  get user samaccountname from user name:   ### 
        $MailBoxPrimarySMTPAddress = "mmckenna@scotcourtstribunals.gov.uk"
        $UserSamAccountName = get-mailbox $($UserNameComboBox1.SelectedItem.ToString()) | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName  
        ## CHECK - continue if only 1 username is in pipe if not exit   ###
        if (($UserSamAccountName | Measure-Object).count -ne 1) {Mainform}
        ##  Run Add Mailbox Permissions Function   ###
        If ($AddFullandSend -eq $true) {
            ##  Check if User already has full access   ###
            $Status = Get-Mailbox mmckenna@scotcourtstribunals.gov.uk | Get-MailboxPermission -User $UserSamAccountName
            If ($Status.AccessRights -eq 'FullAccess') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user ( $($UserNameComboBox1.SelectedItem.ToString()) ) Already has Full Access Permissions to Martin Mckenna mailbox.", "$FormName", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $Form.Close()
                $Form.Dispose()
                Return MainForm
            }
            write-host "is true adding permissions to mailbox"
            Add-MailboxPermission -Identity mmckenna@scotcourtstribunals.gov.uk -User $UserSamAccountName -AccessRights FullAccess -Confirm:$false; Set-Mailbox mmckenna@scotcourtstribunals.gov.uk -GrantSendOnBehalfTo @{add = $UserSamAccountName}
            ##  get current permissions   ###
            $MailboxName = Get-Mailbox -identity mmckenna
            $status = Get-MailboxPermission $MailboxName.Name | Where-Object {$_.AccessRights -eq 'FullAccess'} |  where-object {$_.user -notlike "s-1*" -and @("scs\domain admins", "scs\enterprise admins", "nt authority\system", "scs\organization management") -notcontains $_.User} | Select-Object User, AccessRights
            $StatusTrimmed = $status | Select-Object @{label = 'User'; expression = {$_.User -replace '^SCS\\'}}
            ## confirmation message   ###
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Permissions ADDED to Martin Mckenna mailbox.", "$FormName", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            ## Send confirmation email   ###
            Send-MailMessage -To rmcgaghey@scotcourtstribunals.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "The User ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Full Access permissions ADDED to Martin Mckenna mailbox" -Body " The current users with access permissions on Martins mailbox are:`n`n $($StatusTrimmed.User).`n`n Please Note: Dont reply to this email.`n`nIf you have any queries please email the IT helpdesk." -SmtpServer mail.scotcourts.local
            Send-MailMessage -To rmctaggart@scotcourtstribunals.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "The User ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Full Access permissions ADDED to Martin Mckenna mailbox" -Body " The current users with access permissions on Martins mailbox are:`n`n $($StatusTrimmed.User).`n`n Please Note: Dont reply to this email.`n`nIf you have any queries please email the IT helpdesk." -SmtpServer mail.scotcourts.local
            $Form.Close()
            $Form.Dispose()
            Return MainForm
        }
        ##  Run Remove Mailbox Permissions Function   ###
        If ($RemovePermissions -eq $true) {
            ##  Check if User already has full access   ###
            $Status = Get-Mailbox mmckenna@scotcourtstribunals.gov.uk | Get-MailboxPermission -User $UserSamAccountName
            If ($Status.AccessRights -ne 'FullAccess') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user ( $($UserNameComboBox1.SelectedItem.ToString()) ) does not have Access Permissions to Martin Mckenna mailbox.", "$FormName", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                Remove-Variable Status
                $Form.Close()
                $Form.Dispose()
                Return MainForm
            }
            write-host "Removing Permissions on mailbox"
            Remove-MailboxPermission -Identity mmckenna@scotcourtstribunals.gov.uk -User $UserSamAccountName -AccessRights FullAccess -Confirm:$false; Set-Mailbox mmckenna@scotcourtstribunals.gov.uk -GrantSendOnBehalfTo @{remove = "$UserSamAccountName"}
            ##  get current permissions   ###
            $MailboxName = Get-Mailbox -identity mmckenna
            $status = Get-MailboxPermission $MailboxName.Name | Where-Object {$_.AccessRights -eq 'FullAccess'} |  where-object {$_.user -notlike "s-1*" -and @("scs\domain admins", "scs\enterprise admins", "nt authority\system", "scs\organization management") -notcontains $_.User} | Select-Object User, AccessRights
            $StatusTrimmed = $status | Select-Object @{label = 'User'; expression = {$_.User -replace '^SCS\\'}}
            ## confirmation message   ###
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The user ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Permissions REMOVED from Martin Mckenna mailbox.", "$FormName", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            ## Send confirmation email   ###
            Send-MailMessage -To rmcgaghey@scotcourtstribunals.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "The User ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Access permissions REMOVED from Martin Mckenna mailbox" -Body " The current users with access permissions on Martins mailbox are:`n`n $($StatusTrimmed.User).`n`n Please Note: Dont reply to this email.`n`nIf you have any queries please email the IT helpdesk." -SmtpServer mail.scotcourts.local
            Send-MailMessage -To rmctaggart@scotcourtstribunals.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "The User ( $($UserNameComboBox1.SelectedItem.ToString()) ) has had Access permissions REMOVED from Martin Mckenna mailbox" -Body " The current users with access permissions on Martins mailbox are:`n`n $($StatusTrimmed.User).`n`n Please Note: Dont reply to this email.`n`nIf you have any queries please email the IT helpdesk." -SmtpServer mail.scotcourts.local
            $Form.Close()
            $Form.Dispose()
            Return MainForm
        }   
        $Form.Close()
        $Form.Dispose()
        Return MainForm
    }
}
#################################################################
#########   Completed  -  Create Form1       ####################
#################################################################
#
#################################################################
###########        Create Add Permissions Form   ################
#################################################################

Function AddPermissions {
    $FormName = 'Martin McKenna mailbox - Add Full Access permissions.'
    $Text5 = 'Will have Access and Send On Behalf Of permissions ADDED'
    $TextBox3 = '3. Click Ok to Add mailbox permissions or Cancel.'
    $AddFullandSend = $true
    $RemovePermissions = $false
    Form1
}
#################################################################
########   Completed -   Create Add Full Form         ###########
#################################################################
#
#################################################################
#########        Create Remove Permissions Form     #############
#################################################################
Function RemovePermissions {
    $FormName = 'Martin McKenna mailbox - Remove Mailbox permissions.'
    $Text5 = 'Will have Access and Send On Behalf Of permissions REMOVED'
    $TextBox3 = '3. Click Ok to Remove mailbox permissions or Cancel.'
    $RemovePermissions = $true
    $AddFullandSend = $false
    Form1
}
#################################################################
########      Completed -  Remove Permissions Form    ###########
#################################################################
#
#################################################################
##########        Create Check Permissions Form   ###############
#################################################################
Function CheckPermissions {
    $MailboxName = Get-Mailbox -identity mmckenna
    $status = Get-MailboxPermission $MailboxName.Name | Where-Object {$_.AccessRights -eq 'FullAccess'} |  where-object {$_.user -notlike "s-1*" -and @("scs\domain admins", "scs\enterprise admins", "nt authority\system", "scs\organization management") -notcontains $_.User} |
        Select-Object User, AccessRights
    $StatusTrimmed = $status | Select-Object @{label = 'User'; expression = {$_.User -replace '^SCS\\'}} | Sort-Object user | Out-GridView -Title "List of Users with Full Access permissions on $MailboxName mailbox" -Wait 
    Remove-Variable -Name MailboxName, Status, StatusTrimmed
    Return MainForm
}
#################################################################
########       Completed -  Check Permissions Form    ###########
#################################################################
#
#################################################################
############          Create Main   Form         ################
#################################################################
Function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ##  Set the details of the form   ###
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.width = 750
    $MainForm.height = 500
    $MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $MainForm.MinimizeBox = $False
    $MainForm.MaximizeBox = $False
    $MainForm.FormBorderStyle = 'Fixed3D'
    $MainForm.Text = 'Martin Mckenna Mailbox - Management.'
    $MainForm.Icon = $Icon
    $MainForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    ##  Create a group that will contain the radio buttons   ###
    $MainBox = New-Object System.Windows.Forms.GroupBox
    $MainBox.Location = '40,30'
    $MainBox.size = '650,370'
    $MainBox.text = 'Select an option.'
    ##  Create the radio buttons   ###
    $MainButton1 = New-Object System.Windows.Forms.RadioButton
    $MainButton1.Location = '20,60'
    $MainButton1.size = '600,40'
    $MainButton1.Checked = $true 
    $MainButton1.Text = 'Add - Full Access and Send On Behalf Of permissions for a User.'
    $MainButton2 = New-Object System.Windows.Forms.RadioButton
    $MainButton2.Location = '20,130'
    $MainButton2.size = '600,40'
    $MainButton2.Checked = $false
    $MainButton2.Text = 'Check - current mailbox permissions.'
    $MainButton3 = New-Object System.Windows.Forms.RadioButton
    $MainButton3.Location = '20,200'
    $MainButton3.size = '600,40'
    $MainButton3.Checked = $false
    $MainButton3.Text = 'Remove - Full Access and Send On Behalf Of permissions for a User.'
    ##  Add an OK button   ###
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '500,415'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    ##  Add a cancel button   ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '375,415'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $MainForm.Close()
            $MainForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
    ##  Add all the Form controls on one line   ### 
    $MainForm.Controls.AddRange(@($MainBox, $OKButton, $CancelButton))
    ##  Add all the GroupBox controls on one line   ###
    $MainBox.Controls.AddRange(@($MainButton1, $MainButton2, $MainButton3, $MainButton4, $MainButton5, $MainButton6))
    ##  Assign the Accept and Cancel options in the form to the corresponding buttons   ###
    $MainForm.AcceptButton = $OKButton
    $MainForm.CancelButton = $CancelButton
    ##  Activate the form   ###
    $MainForm.Add_Shown( {$MainForm.Activate()})    
    ##  Get the results from the button click   ###
    $dialogResult = $MainForm.ShowDialog()
    ##  If the OK button is selected   ###
    if ($dialogResult -eq 'OK') {
        ##  Check the current state of each radio button and respond   ###
        if ($MainButton1.Checked) {AddPermissions}
        elseif ($MainButton2.Checked) {CheckPermissions}
        elseif ($MainButton3.Checked) {RemovePermissions}
    }
}
Mainform