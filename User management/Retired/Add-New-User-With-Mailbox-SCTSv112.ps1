﻿# Add New SCTS User Account with Exchange 2013 mailbox:
# Author         John Mckay
# Date           13/02/2018
# Version        1.12
# Purpose        To add a new user with an exchange 2013 mailbox.A log is kept in \\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox\NewSCTSUserLog.txt
# Usage          Helpdesk staff run a shorcut to Add-New-User-With-Mailbox-SCTS.exe in \\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox 
# Changes        v1.12 JM changed send email to helpdesk to use $env:UserName instead of helpdesk.
#                v1.11 JM enabled import exchange module line 23  & 24 was disabled by mistake.
#                Added start transcript to get log file.
#                Added line 34 to hide command window when script runs.
#                v1.10 JM Changed to use v2 form to keep uniformity with other helpdesk forms.
#                v1.03 JM Change to "Add Full Control permissions for user on p drive folder" line 584 added 
#                "Add user as owner on p drive folder" $acl.SetOwner([System.Security.Principal.NTAccount]"$LogOnName")
#                to set user as folder owner. 
#                and removed $ruleuser = New-Object System.Security.AccessControl.FileSystemAccessRule
#                v1.02 JM Added "Disable Pop, OWA, Imap & ActiveSync for user" line 595 to disable unused email protocols.
#                v1.01 JM Change to "Create mailbox for userAdd" line 590 to Enable-MailBox -Identity $LogonName@scotcourts.local to allow
#                mailbox to be created for user with same displayname 
#
Start-Transcript -Path "\\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox\Logs\NewSCTSUserlog.txt" -append
########################################################################
#######         Create Session with Exchange 2013         ##############
########################################################################
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
#
#########################################################################
############     Set icon for all forms and subforms      ###############
#########################################################################
$Icon = "\\saufs01\IT\Enterprise Team\Usermanagement\icons\user.ico"
#########################################################################
#######              Show Start Message:                  ###############
#########################################################################
Add-Type -AssemblyName System.Windows.Forms 
[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("This script creates a New User Account with a mailbox in Exchange 2013.`n`nThe New Account will be created in the NewUsers OU in AD & the account will be disabled.`n`nBefore use the New User Account needs to be moved from the NewUsers OU to the correct User OU in AD & enabled.`n`nPlease click OK to continue or Cancel to exit", "Add New SCTS User Account with Exchange 2013 mailbox.", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
if ($StartMessage -eq 'Cancel') { exit } 
else {
######################################################################
####     Create SubForm  Add New Shared mailbox Sub Form          ####
######################################################################
Function AddNewUser{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    ### Set the details of the form. ###
    $NewUserForm1 = New-Object System.Windows.Forms.Form
    $NewUserForm1.width = 780
    $NewUserForm1.height = 760
    $NewUserForm1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $NewUserForm1.Controlbox = $false
    $NewUserForm1.Icon = $Icon
    $NewUserForm1.FormBorderStyle = "Fixed3D"
    $NewUserForm1.Text = "Add New SCTS User Account with Exchange 2013 mailbox v1.12"
    $NewUserForm1.Font = New-Object System.Drawing.Font("Ariel",10)
    ### Create group 1 box in form. ####
    $NewUserBox1 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox1.Location = '40,30'
    $NewUserBox1.size = '700,200'
    $NewUserBox1.text = "1. Enter the new Users details (Mandatory fields are marked with *M*) :"
    ### Create group 1 box text labels. ###
    $NewUsertextLabel1 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel1.Location = '20,35'
    $NewUsertextLabel1.size = '350,40'
    $NewUsertextLabel1.Text = "First Name :  (e.g. Joseph with capital 'J'). *M*" 
    $NewUsertextLabel2 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel2.Location = '20,75'
    $NewUsertextLabel2.size = '350,40'
    $NewUsertextLabel2.Text = "Last Name :  (e.g. Bloggs with capital 'B'). *M*" 
    $NewUsertextLabel3 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel3.Location = '20,112'
    $NewUsertextLabel3.size = '370,40'
    $NewUsertextLabel3.Text = "LogOnName :  (e.g jbloggs,sheriffjbloggs,ladybloggs). *M*" 
    $NewUsertextLabel4 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel4.Location = '20,150'
    $NewUsertextLabel4.size = '370,40'
    $NewUsertextLabel4.Text = "Account Expiry Date : (e.g 25/03/2018)." 
    ### Create group 1 box text boxes. ###
    ## First name textbox ##
    $NewUsertextBox1 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox1.Location = '425,30'
    $NewUsertextBox1.Size = '250,40'
    $NewUsertextBox1.add_TextChanged({$NewUsertextLabel9.Text = "$($NewUsertextBox1.text)"})
    $NewUsertextBox1.Add_TextChanged({If($This.Text -and $NewUsertextBox2.Text -and $NewUsertextBox3.Text){$ContinueButton.Enabled = $True}Else{$ContinueButton.Enabled = $False}})  
    ## Last name textbox ##
    $NewUsertextBox2 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox2.Location = '425,70'
    $NewUsertextBox2.Size = '250,40'
    $NewUsertextBox2.add_textChanged({$NewUsertextLabel10.Text = "$($NewUsertextBox2.text)"})
    $NewUsertextBox2.Add_TextChanged({If($This.Text -and $NewUsertextBox1.Text -and $NewUsertextBox3.Text){$ContinueButton.Enabled = $True}Else{$ContinueButton.Enabled = $False}}) 
    ## LogOnName textbox ##
    $NewUsertextBox3 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox3.Location = '425,105'
    $NewUsertextBox3.Size = '250,40'
    $NewUsertextBox3.add_TextChanged({$NewUsertextLabel8.Text = "$($NewUsertextBox3.text)"})
    $NewUsertextBox3.Add_TextChanged({If($This.Text -and $NewUsertextBox1.Text -and $NewUsertextBox2.Text){$ContinueButton.Enabled = $True}Else{$ContinueButton.Enabled = $False}})  
    ## Acc Expire textbox ##
    $NewUsertextBox4 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox4.Location = '425,145'
    $NewUsertextBox4.Size = '250,40'
    $NewUsertextBox4.add_textChanged({$AccountExpires = "$($NewUsertextBox4.text)"})
    ### Create group 2 box in form. ###
    $NewUserBox2 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox2.Location = '40,240'
    $NewUserBox2.size = '700,240'
    $NewUserBox2.text = "2. Enter the new users AD details:"
    ### Create group 2 box text labels. ###
    $NewUser2textLabel1 = New-Object “System.Windows.Forms.Label”;
    $NewUser2textLabel1.Location = '20,35'
    $NewUser2textLabel1.size = '350,40'
    $NewUser2textLabel1.Text = "Office field in AD:" 
    $NewUser2textLabel2 = New-Object “System.Windows.Forms.Label”;
    $NewUser2textLabel2.Location = '20,75'
    $NewUser2textLabel2.size = '350,40'
    $NewUser2textLabel2.Text = "Description field in AD:" 
    $NewUser2textLabel3 = New-Object “System.Windows.Forms.Label”;
    $NewUser2textLabel3.Location = '20,115'
    $NewUser2textLabel3.size = '350,40'
    $NewUser2textLabel3.Text = "Distribution List field in AD:" 
    $NewUser2textLabel4 = New-Object “System.Windows.Forms.Label”;
    $NewUser2textLabel4.Location = '20,155'
    $NewUser2textLabel4.size = '370,40'
    $NewUser2textLabel4.Text = "Security Group field in AD (to access p and s drives):" 
    $NewUser2textLabel5 = New-Object “System.Windows.Forms.Label”;
    $NewUser2textLabel5.Location = '20,195'
    $NewUser2textLabel5.size = '370,40'
    $NewUser2textLabel5.Text = "LogOn Script field in AD:"
    ####################################################################
    #############    Define inputs for combo boxes     #################
    ####################################################################
    $OfficeList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\Office.csv"
    $SecurityGroupsList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\SecurityGroups.csv"
    $LogOnScriptList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\LogOnScript.csv"
    $DescriptionList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\description.csv"
    ### Create group 2 box combo boxes. ###
    ###  Populate "Office" ComboBox1   ###
    $NewUser2comboBox1 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox1.Location = '425,30'
    $NewUser2comboBox1.Size = '250,40'
    $NewUser2comboBox1.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox1.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox1.Sorted = $false;
    $NewUser2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox1.SelectedItem = $NewUser2comboBox1.Items[0]
    $NewUser2comboBox1.DataSource = $OfficeList.Office 
    $NewUser2comboBox1.add_SelectedIndexChanged({$NewUser2OfficeSelect.Text = "$($NewUser2comboBox1.SelectedItem.ToString())"})
    ###  Populate "Description" ComboBox2 ###
    $NewUser2comboBox2 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox2.Location = '425,70'
    $NewUser2comboBox2.Size = '250,40'
    $NewUser2comboBox2.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox2.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox2.Sorted = $false;
    $NewUser2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox2.SelectedItem = $NewUser2comboBox2.Items[0]
    $NewUser2comboBox2.DataSource = $DescriptionList.Description 
    $NewUser2comboBox2.add_SelectedIndexChanged({$NewUser2DescriptionSelect.Text = "$($NewUser2comboBox2.SelectedItem.ToString())"})
    ###  Populate "Distribution List" ComboBox3 ###
    $NewUser2comboBox3 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox3.Location = '425,110'
    $NewUser2comboBox3.Size = '250,40'
    $NewUser2comboBox3.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox3.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox3.Sorted = $true;
    $NewUser2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox3.SelectedItem = $NewUser2comboBox3.Items[0]
    $NewUser2comboBox3.DataSource = Get-ADObject –Filter * -Searchbase "ou=distribution lists,ou=groups,ou=courts,dc=scotcourts,dc=local" | Select-Object name | Select-Object -ExpandProperty Name  
    $NewUser2comboBox3.add_SelectedIndexChanged({$NewUser2DistributionSelect.Text = "$($NewUser2comboBox3.SelectedItem.ToString())"})
    ###  Populate "Security Group" ComboBox4 ###
    $NewUser2comboBox4 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox4.Location = '425,150'
    $NewUser2comboBox4.Size = '250,40'
    $NewUser2comboBox4.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox4.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox4.Sorted = $false;
    $NewUser2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox4.SelectedItem = $NewUser2comboBox4.Items[0]
    $NewUser2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
    $NewUser2comboBox4.add_SelectedIndexChanged({$NewUser2SecuritySelect.Text = "$($NewUser2comboBox4.SelectedItem.ToString())"})
    ###  Populate "LogOn Script" ComboBox5 ###
    $NewUser2comboBox5 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox5.Location = '425,190'
    $NewUser2comboBox5.Size = '250,40'
    $NewUser2comboBox5.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox5.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox5.Sorted = $false;
    $NewUser2comboBox5.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox5.SelectedItem = $NewUser2comboBox5.Items[0]
    $NewUser2comboBox5.DataSource = $LogOnScriptList.Logonscript 
    $NewUser2comboBox5.add_SelectedIndexChanged({$NewUser2LogonSelect.Text = "$($NewUser2comboBox5.SelectedItem.ToString())"})
    ### Create group 2 labels to take combobox output ###
    $NewUser2OfficeSelect = New-Object “System.Windows.Forms.Label”;
    $NewUser2OfficeSelect.Location = '20,600'
    $NewUser2OfficeSelect.size = '350,40'
    $NewUser2DescriptionSelect = New-Object “System.Windows.Forms.Label”;
    $NewUser2DescriptionSelect.Location = '20,650'
    $NewUser2DescriptionSelect.size = '350,40'
    $NewUser2DistributionSelect = New-Object “System.Windows.Forms.Label”;
    $NewUser2DistributionSelect.Location = '20,700'
    $NewUser2DistributionSelect.size = '350,40'
    $NewUser2SecuritySelect = New-Object “System.Windows.Forms.Label”;
    $NewUser2SecuritySelect.Location = '20,750'
    $NewUser2SecuritySelect.size = '350,40'
    $NewUser2LogonSelect = New-Object “System.Windows.Forms.Label”;
    $NewUser2LogonSelect.Location = '20,800'
    $NewUser2LogonSelect.size = '350,40'
    ### Create group 3 box in form. ###
    $NewUserBox3 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox3.Location = '40,495'
    $NewUserBox3.size = '700,115'
    $NewUserBox3.text = "3. Check the details below are correct before proceeding:"
    ### Create group 3 box text labels.
    ## message label ##
    $NewUsertextLabel5 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel5.Location = '20,40'
    $NewUsertextLabel5.size = '350,30'
    $NewUsertextLabel5.Text = "New User will appear in AD and Global Address List as:" 
    ## DisplayName label ##
    $NewUsertextLabel6 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel6.Location = '20,75'
    $NewUsertextLabel6.Size = '100,30'
    $NewUsertextLabel6.ForeColor = "Blue"
    $NewUsertextLabel6.text = $DisplayName
    ## message label ##
    $NewUsertextLabel7 = New-Object “System.Windows.Forms.Label”;
    $NewUsertextLabel7.Location = '430,40'
    $NewUsertextLabel7.size = '200,30'
    $NewUsertextLabel7.Text = "With the LogOnName:"
    ## LogOnName label ##
    $NewUsertextLabel8 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel8.Location = '460,75'
    $NewUsertextLabel8.Size = '400,30'
    $NewUsertextLabel8.ForeColor = "Blue"
    ## First name label ##
    $NewUsertextLabel9 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel9.Location = '20,300'
    $NewUsertextLabel9.Size = '250,30'
    $NewUsertextLabel9.ForeColor = "Blue"
    $NewUsertextLabel9.add_TextChanged({$NewUsertextLabel6.Text = "$($NewUsertextLabel10.text + ", " + $NewUsertextLabel9.text)"})
    ## Last name label ##
    $NewUsertextLabel10 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel10.Location = '20,350'
    $NewUsertextLabel10.Size = '250,30'
    $NewUsertextLabel10.ForeColor = "Blue"
    $NewUsertextLabel10.add_TextChanged({$NewUsertextLabel6.Text = "$($NewUsertextLabel10.text + ", " + $NewUsertextLabel9.text)"})
    ### Create group 4 box in form. ###
    $NewUserBox4 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox4.Location = '40,620'
    $NewUserBox4.size = '700,40'
    $NewUserBox4.text = "4. Click Continue or Exit:"
    $NewUserBox4.button
    ### Add an OK button ###
    $ContinueButton = new-object System.Windows.Forms.Button
    $ContinueButton.Location = '640,675'
    $ContinueButton.Size = '100,40'          
    $ContinueButton.Text = 'Continue'
    $ContinueButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,675'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Exit"
    $CancelButton.add_Click({
    $NewUserForm1.Close()
    $NewUserForm1.Dispose()
    $form1.[System.Environment]::Exit(0)})
    ### Add all the Form controls ### 
    $NewUserForm1.Controls.AddRange(@($NewUserBox1,$NewUserBox2,$NewUserBox3,$NewUserBox4,$ContinueButton,$CancelButton))
    #### Add all the GroupBox controls ###
    $NewUserBox1.Controls.AddRange(@($NewUsertextLabel1,$NewUsertextLabel2,$NewUsertextLabel3,$NewUsertextLabel4,$NewUsertextLabel5,$NewUsertextLabel6,$NewUsertextLabel7,$NewUsertextLabel8,$NewUsertextLabel9,$NewUsertextLabel10,$NewUsertextBox1,$NewUsertextBox2,$NewUsertextBox3,$NewUsertextBox4))
    $NewUserBox2.Controls.AddRange(@($NewUser2textLabel1,$NewUser2textLabel2,$NewUser2textLabel3,$NewUser2textLabel4,$NewUser2textLabel5,$NewUser2comboBox1,$NewUser2comboBox2,$NewUser2comboBox3,$NewUser2comboBox4,$NewUser2comboBox5,$NewUser2OfficeSelect,$NewUser2DescriptionSelect,$NewUser2DistributionSelect,$NewUser2SecuritySelect,$NewUser2LogonSelect))
    $NewUserBox3.Controls.AddRange(@($NewUsertextLabel5,$NewUsertextLabel6,$NewUsertextLabel7,$NewUsertextLabel8,$NewUsertextLabel9,$NewUsertextLabel10))
    #### Activate the form ###
    $NewUserForm1.Add_Shown({$NewUserForm1.Activate()})    
    $dialogResult = $NewUserForm1.ShowDialog()
    #####################################################################
    ########                    set variables               ############# 
    #####################################################################
    $FirstName = $NewUsertextBox1.text
    $LastName = $NewUsertextBox2.text
    $LogOnName = $NewUsertextBox3.text
    $AccountExpires = $NewUsertextBox4.text
    $DisplayName = $LastName + ", " + $FirstName
    #####################################################################
    ########   Don't accept null username or mailbox     ################ 
    #####################################################################
    if ($NewUsertextBox1.text -eq "") {[System.Windows.Forms.MessageBox]::Show("You need to type in details !!!!!`n`nTrying to enter blank fields is never a good idea.", "Add New Shared Mailbox", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    $NewUserForm1.Close()
    $NewUserForm1.Dispose()
    break}
    #########################################################################
    ##########  Check to see if Samaccountname is already in use  ###########
    #########################################################################
    $User = Get-ADUser -Filter {sAMAccountName -eq $LogOnName}
    If ($User -ne $Null) {
    Add-Type -AssemblyName System.Windows.Forms 
    [System.Windows.Forms.MessageBox]::Show("The LogOnName $LogOnName can't be used because it's assigned to an existing user account.`n`nThe next page will display the current usernames in use for $LogOnName`n`nPlease use a LogOnName that's not currently in use.", "ERROR - CAN'T ADD NEW USER", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    Get-AdUser -Filter "SamAccountName -like '$LogOnName*'" | Select-Object SamAccountName | Out-GridView -title "User accounts currently in use"
    $NewUserForm1.Close()
    $NewUserForm1.Dispose()
    Remove-Variable DisplayName
    AddNewUser} 
    Else {
    #####################################################################
    ##   CHECK - continue if only 1 EmailName in pipe if not exit      ##
    #####################################################################
    if (($LogOnName | Measure-Object).count -ne 1) {AddNewUser}
    #####################################################################
    #
    $Password = "Helpdesk123"
    $Office =  $NewUser2OfficeSelect.text
    $Description = $NewUser2DescriptionSelect.Text
    $DistributionGroup = $NewUser2DistributionSelect.Text
    $SecurityGroup = $NewUser2SecuritySelect.Text
    $LogOnScript = $NewUser2LogonSelect.Text
    #
    ##########################################################################
    #############        Create AD account         ###########################
    ##########################################################################
    New-AdUser  -Name "$DisplayName" -SamAccountName $LogonName -Path "OU=NewUsers,DC=scotcourts,DC=local" –AccountPassword ($Password | ConvertTo-SecureString -AsPlainText –Force)
    ##########################################################################
    #
    ##########################################################################
    #############  Create form to pause for 5 sec   ##########################
    ##########################################################################
    Add-Type -AssemblyName System.Windows.Forms
    ### Build Form ###
    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = "Add New User Account"
    $objForm.Size = New-Object System.Drawing.Size(450,270)
    $objForm.StartPosition = "CenterScreen"
    $objForm.Controlbox = $false
    # Add Label
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(80,50) 
    $objLabel.Size = New-Object System.Drawing.Size(300,120)
    $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the NewUser OU.`n`nPlease Wait. .............."
    $objForm.Controls.Add($objLabel)
    # Show the form
    $objForm.Show()| Out-Null
    # wait 10 seconds
    Start-Sleep -Seconds 5
    # destroy form
    $objForm.Close() | Out-Null
    #
    ##########################################################################
    #######    Add security & distribution group permissions    ############## 
    ##########################################################################
    Add-ADGroupMember -PassThru "DomainShareAccess" $LogOnName
    Add-ADGroupMember -PassThru "CN=$Securitygroup,OU=Security Groups,OU=SCS Users,DC=scotcourts,DC=local" $LogOnName 
    Add-ADGroupMember -PassThru "$DistributionGroup" $LogOnName
    ##########################################################################
    ###################     Set user AD properties             ############### 
    ##########################################################################
    Set-AdUser –PassThru -Identity $LogonName –GivenName "$FirstName" –Surname "$LastName" -DisplayName "$Displayname" 
    Set-ADUser –PassThru -Identity $LogonName -Office $Office –Description $Description 
    Set-AdUser –PassThru -Identity $LogonName -UserPrincipalName "$LogonName@scotcourts.local"
    Set-ADUser –PassThru -Identity $LogonName –HomeDrive "P:" -HomeDirectory "\\scotcourts.local\data\users\$LogOnName"
    Set-ADUser –PassThru -Identity $LogonName -ProfilePath "\\scotcourts.local\data\profiles\$LogOnName"
    Set-ADUser –PassThru -Identity $LogonName -AccountExpirationDate $AccountExpires
    Set-ADUser –PassThru -Identity $LogonName -ScriptPath $LogOnScript
    ##########################################################################
    ######## Set password change at next logon      ##########################
    ##########################################################################
    #Set-ADUser -Identity $LogonName -ChangePasswordAtLogon $true
    ##########################################################################
    #############       Disable New User account    ##########################
    ##########################################################################
    Set-ADUser -Identity $LogonName  -Enabled $False 
    ##########################################################################
    #############  Create form to pause for 5 sec  ##########################
    ##########################################################################
    Add-Type -AssemblyName System.Windows.Forms
    # Build Form
    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = "Add New User Account"
    $objForm.Size = New-Object System.Drawing.Size(450,270)
    $objForm.StartPosition = "CenterScreen"
    $objForm.Controlbox = $false
    # Add Label
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(80,50) 
    $objLabel.Size = New-Object System.Drawing.Size(300,120)
    $objLabel.Text = "A New User Mailbox is being created in Exchange 2013 with the details you entered. `n`n`n`nPlease Wait. .............."
    $objForm.Controls.Add($objLabel)
    # Show the form
    $objForm.Show()| Out-Null
    # wait 10 seconds
    Start-Sleep -Seconds 5
    # destroy form
    $objForm.Close() | Out-Null
    #
    ##########################################################################
    ###########     Add users P drive folder on saufs01    ###################
    ##########################################################################
    New-Item -Path \\saufs01\users -Name $LogonName -ItemType Directory -Force
    ##########################################################################
    #####   Add user as owner on p drive & set to inherit permissions    #####
    ##########################################################################
    $acl = Get-Acl \\saufs01\users\$LogOnName
    $acl.SetAccessRuleProtection($false, $false)
    $acl.SetOwner([System.Security.Principal.NTAccount]"$LogOnName")
    Set-Acl \\saufs01\users\$LogOnName $acl
    ##########################################################################
    ###########         Set permissions complete           ###################
    ##########################################################################
    #
    ##########################################################################
    ###########        Create mailbox for user          ######################
    ##########################################################################
    Enable-MailBox -Identity $LogonName@scotcourts.local
    #
    ##########################################################################
    #######   Disable Pop, OWA, Imap & ActiveSync for user ###################
    ##########################################################################
    Set-CASMailbox -Identity $LogonName -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
    ##########################################################################
    ###########        Generate Form complete           ######################
    ##########################################################################
    Add-Type -AssemblyName System.Windows.Forms 
    $StartMessage = [System.Windows.Forms.MessageBox]::Show("The User account and mailbox have been created in the NewUsers OU.`n`nNote1:  The user account needs to be moved from the NewUsers OU to the correct user OU in AD.:`n`nNote2:  The user account needs to be enabled before use.", "New User Account", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    ##########################################################################
    ###########      Send email to helpdesk             ######################
    ##########################################################################
    Send-MailMessage -To helpdesk@scotcourts.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New User Account $LogOnName added. The new User needs moved out of the New User OU in AD." -Body "A new user account has been added:`n`nUserName:   $DisplayName`n`nLogOnName:   $LogOnName`n`nLocation:   $Description`n`nDistribution List:  $DistributionGroup`n`nSecurity Group:   $SecurityGroup " -SmtpServer mail.scotcourts.local
    Remove-Variable DisplayName
    AddNewUser}
    }}
    AddNewUser

# SIG # Begin signature block
# MIIQEwYJKoZIhvcNAQcCoIIQBDCCEAACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUjKHIHWK20psqf+PYKtnt/5D7
# sLuggg19MIIGYDCCBUigAwIBAgITbwAAACIK+5RtwBxYlgAAAAAAIjANBgkqhkiG
# 9w0BAQsFADBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxGjAYBgoJkiaJk/IsZAEZ
# FgpzY290Y291cnRzMRMwEQYDVQQDEwpTQ1RTLUVudENBMB4XDTE3MTAxOTEzMTA0
# NFoXDTE4MTAxOTEzMTA0NFowezEVMBMGCgmSJomT8ixkARkWBWxvY2FsMRowGAYK
# CZImiZPyLGQBGRYKc2NvdGNvdXJ0czEZMBcGA1UECxMQU2VydmljZSBBY2NvdW50
# czErMCkGA1UEAxMiR2l0TGFiIE11bHRpUnVubmVyIHNlcnZpY2UgYWNjb3VudDCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMr0G3ypKLBZjArZBdaOjhNO
# Hl1r2itu26RsLJOJA1G6PDps3ll6pU+YOPi1VFVKGGVcdaEPEGJCD3oirHN2KkB9
# BSk7j8UAsR6ef9v6/SoQZdrKV2oNaqlCgbeVVtJ0JfHh5AQLYSlw4CRmpRuPWODT
# TXmFoLDvMYU2Ee4NF4hLthHXdiKXfBKiocgM79ZOCBHpMpXjZ0VXcNbX1kijgtMR
# PDlvw9oaCXge8ySqriMCMczSccw1xyQsnwuktwV+Y6vt14/Q4r7Bkwxob/oHtayn
# K1OYlnvNt6pWLEO9bP7ZE3FTq+oqaWGSosiEToQs+m6yi8lcnBnj7+6zneyfE+cC
# AwEAAaOCAw4wggMKMD4GCSsGAQQBgjcVBwQxMC8GJysGAQQBgjcVCITTjy6GoKId
# gb2FLYLKwHSBjudwgTKCvJdkgru9MQIBZAIBBTATBgNVHSUEDDAKBggrBgEFBQcD
# AzAOBgNVHQ8BAf8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAd
# BgNVHQ4EFgQU206JDOCO+3ydDvD+jYypHAU6EI4wHwYDVR0jBBgwFoAUDutbzfmF
# SD3EskvYrYce2x6UvDgwgf0GA1UdHwSB9TCB8jCB76CB7KCB6YaBtmxkYXA6Ly8v
# Q049U0NUUy1FbnRDQSxDTj1TQ1RDQTAzLENOPUNEUCxDTj1QdWJsaWMlMjBLZXkl
# MjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPXNjb3Rj
# b3VydHMsREM9bG9jYWw/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29i
# amVjdENsYXNzPWNSTERpc3RyaWJ1dGlvblBvaW50hi5odHRwOi8vcGtpLnNjb3Rj
# b3VydHMubG9jYWwvcGtpL1NDVFMtRW50Q0EuY3JsMIIBBQYIKwYBBQUHAQEEgfgw
# gfUwQgYIKwYBBQUHMAKGNmh0dHA6Ly9wa2kuc2NvdGNvdXJ0cy5sb2NhbC9wa2kv
# U0NUQ0EwM19TQ1RTLUVudENBLmNydDCBrgYIKwYBBQUHMAKGgaFsZGFwOi8vL0NO
# PVNDVFMtRW50Q0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENO
# PVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9c2NvdGNvdXJ0cyxEQz1sb2Nh
# bD9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1
# dGhvcml0eTA9BgNVHREENjA0oDIGCisGAQQBgjcUAgOgJAwic3ZjX0dMTXVsdGlS
# dW5uZXJAc2NvdGNvdXJ0cy5sb2NhbDANBgkqhkiG9w0BAQsFAAOCAQEAvM+SpIar
# oya99NYp6/Oh4GnTzg4SCJbmqpZ8NnMBXiXDxvG0CvlRcXH2VjsvYQO8Q5AVPBPf
# uoysgqP0YhujmnyIsSfZPqgF07V/y7ymiyeS7vHb/7Qaz8FrfCduHKFJUtkuFJ0W
# 7UI6IKhcTo47cTqdDabub32Dxh6+6Q38q9yabW/GYdvqyobKuYb7/l5Nx86Bh28S
# gsQ0g+MlMskBYd8nePr3lhdwiFY45kTp8mbp8EXIjmcvyeNbaMJt7P8/oKOt2Oo4
# XguGe+s8hV6Ss+ok246SYFFQXQXvmbUT8GBhOw5k8z5PjVPywQ206QXHKMXIfe/N
# GU1CmRvk+WXNqjCCBxUwggT9oAMCAQICExwAAAADZ7b0KwMLG2wAAAAAAAMwDQYJ
# KoZIhvcNAQELBQAwFjEUMBIGA1UEAxMLU0NUUy1Sb290Q0EwHhcNMTcwNTExMTIw
# MTQ3WhcNMjIwNTExMTAwODU1WjBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxGjAY
# BgoJkiaJk/IsZAEZFgpzY290Y291cnRzMRMwEQYDVQQDEwpTQ1RTLUVudENBMIIB
# IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzIFPQH48gwZSTTQMh5lZUkfK
# mIXSCZsSRhJGzXPTQ68ZvO3xTGC5qtU0eAN3mJCHTC4aYEM7bOt1958P+3BVd3YJ
# +7eIlzqYxd4hD3SguYLlzvLxsk5ZNZC+LSEmLa1ANZ5oBtB0BG+ZW0qQmXReI+Fi
# qaimRQZvqaXxSq7EpQG7oyZjZnh7/WdzQF2RsRKIk2Mj2L6GjFPRTqK7brkFjygt
# XxQj6S7nmuj3+1WJA/osHGtFtgvqWckTz5F2xxcDMDMeMUFUp0WKY2mxAfe5KSr0
# lu1QUQ8e4oQDH1vfBGmMejba7JNb6jN1gGYI7FED16iHsDyRfjEOZQdSVnZMnwID
# AQABo4IDKDCCAyQwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFA7rW835hUg9
# xLJL2K2HHtselLw4MIGJBgNVHSAEgYEwfzB9BggqAwSrM0JOCTBxMDoGCCsGAQUF
# BwICMC4eLABMAGUAZwBhAGwAIABQAG8AbABpAGMAeQAgAFMAdABhAHQAZQBtAGUA
# bgB0MDMGCCsGAQUFBwIBFidodHRwOi8vcGtpLnNjb3Rjb3VydHMubG9jYWwvcGtp
# L2Nwcy50eHQwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGG
# MA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAUxnH46FF7yQz3EOXV8gYZRqoK
# EAgwgf8GA1UdHwSB9zCB9DCB8aCB7qCB64YvaHR0cDovL3BraS5zY290Y291cnRz
# LmxvY2FsL3BraS9TQ1RTLVJvb3RDQS5jcmyGgbdsZGFwOi8vL0NOPVNDVFMtUm9v
# dENBLENOPVNDVENBMDEsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2Vz
# LENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9c2NvdGNvdXJ0cyxEQz1s
# b2NhbD9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9
# Y1JMRGlzdHJpYnV0aW9uUG9pbnQwggEHBggrBgEFBQcBAQSB+jCB9zBDBggrBgEF
# BQcwAoY3aHR0cDovL3BraS5zY290Y291cnRzLmxvY2FsL3BraS9TQ1RDQTAxX1ND
# VFMtUm9vdENBLmNydDCBrwYIKwYBBQUHMAKGgaJsZGFwOi8vL0NOPVNDVFMtUm9v
# dENBLENOPUFJQSxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNl
# cyxDTj1Db25maWd1cmF0aW9uLERDPXNjb3Rjb3VydHMsREM9bG9jYWw/Y0FDZXJ0
# aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRob3JpdHkw
# DQYJKoZIhvcNAQELBQADggIBAB4huV6ivXiA0nERPJjwGtQsQB1sgb5mPgFPwpG6
# SUzPk1VvI+PjbRIggUu9HsWrn/aDpXzEGBquRHa+m1g86EPq3zcjR0zBwtTUCcUb
# k1F0zRxSVL5coLSDCPKHeJFu4M5dqsIrdo9qiZ6bjknxMkWwUIgssrFR4F1q/5Jm
# zNwsoZN5BGfZEFqzlCWZrr3uqOA5vnsXOp95fXpnwPyBTQFYU93edggWS75rtm4d
# kTLKSapfjNO/GUGGo0LDrRFiqD71rnlJoaZBEynmpBAVOqXjW3YDkEicpM7mCHT8
# 9qZXGVWLt3NfX5q1d0IfmU6pkS0dB+FbyNZzsvduze2jsEkQSLXFjo25qQP/LhZF
# cajsglBvXRCa8WrzBDbU/RC3jZeU0yXI+wagwxKVIM8r3HMjnvWYRDavO7FvFSpB
# 9qAgxJFuItmwonm8lp+iF/vCKSwCVng7KOzZvI/vsftkPZ2B8QcV8CR453nJ+W60
# z4Jv8hSbW65QtoWWgq1DddhSPxtHA6aWoG+Ch9/ehVgv7Idi5NP/BafEqs9HTM7a
# 94u/hzFVPalOsAjUyYHOzMjhCat7d/LOvjadyPW2+jr47dlX53YHAc3DpOpcmemQ
# 6HU3nquAuMPml4J1v93nSh9ldkpnDKpdhJ5X10C8agedhwZSf0YmZ2Fnv2IN+Vo6
# 8KBDMYICADCCAfwCAQEwXzBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxGjAYBgoJ
# kiaJk/IsZAEZFgpzY290Y291cnRzMRMwEQYDVQQDEwpTQ1RTLUVudENBAhNvAAAA
# Igr7lG3AHFiWAAAAAAAiMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTCu6cAhsVvH9QT531mW+du
# YM920TANBgkqhkiG9w0BAQEFAASCAQADjVQaF7jWetEiFiFY0FT6tVyDVRlZYu0h
# MjrIEXjSY+BbasOFSNtbjsjI7SXs7mAGUcckp5y4AI6bmcDw1iUUpapDxw6TmgVr
# KxMZeWfjTpPA2hxB23GofXLuVSnU3EaOlcSdUwcZLv5yd/1gLwDFBsAhixoFx4UD
# tvpwWakBAH0YOTEPnk1NrcWBwx+dyX5QlMzi8WYOsSiu2vLYIy8IlVmubnfaGDiX
# 7Gtu0qcqy26fvwyw2fgIdHe5tDsIdHRCecqO0nNLFTfd/SIHtzO4MlLWAW9jR8Yt
# Yq+TUtbVI1s5hcvPGUM6R9tXK4V2e0u2XEHVdA0r/Yyoa+wfjpLY
# SIG # End signature block
