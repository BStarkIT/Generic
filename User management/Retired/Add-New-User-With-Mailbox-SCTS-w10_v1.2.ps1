# Add New SCTS User Account with Exchange 2013 mailbox:
# Author         John Mckay
# Date           13/02/2018
# Version        1.2 (1.13)
# Purpose        To add a new user with an exchange 2013 mailbox.A log is kept in \\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox\NewSCTSUserLog.txt
# Usage          Helpdesk staff run a shorcut to Add-New-User-With-Mailbox-SCTS.exe in \\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox 
# Changes        v1.2  GK 8/1/20 updated all unc paths to use dsf format.
#                                uses unified w10 user OU for all new users accounts.
#                                references to AD home folder and ADvprofile path removed.
#                                new transcript path.
#                v1.13(1.1) GK change to take account of the new personal folder location.
#                v1.12 JM changed send email to helpdesk to use $env:UserName instead of helpdesk.
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
Start-Transcript -Path "\\saufs01\it\Enterprise Team\UserManagement\Add-New-User-With-Mailbox\Logs\NewSCTSUserlog-w10.txt" -append

#######         Create Session with Exchange 2013         ##############
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$Version='1.2'


############     Set icon for all forms and subforms      ###############
$Icon = "\\saufs01\it\Enterprise Team\Usermanagement\icons\user.ico"

#######              Show Start Message:                  ###############
Add-Type -AssemblyName System.Windows.Forms 
[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("This script creates a New User Account with a mailbox in Exchange 2013.`n`nThe New Account will be created in the 'SCTS\User Accounts\SCTS Users' OU in AD & the account will be disabled.`n`nBefore use the New User Account needs to be  enabled.`n`nPlease click OK to continue or Cancel to exit", "Add New SCTS User Account with Exchange 2013 mailbox.",
    [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
if ($StartMessage -eq 'Cancel') {
     exit 
} 
else {

    ##     Create SubForm  Add New Shared mailbox Sub Form          ####
    function AddNewUser{
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
        $NewUserForm1.Text = "Add New SCTS User Account with Exchange 2013 mailbox (W10) v$Version"
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
        #region Remove Lines
        #$NewUser2textLabel5 = New-Object “System.Windows.Forms.Label”;
        #$NewUser2textLabel5.Location = '20,195'
        #$NewUser2textLabel5.size = '370,40'
        #$NewUser2textLabel5.Text = "LogOn Script field in AD:"
        #endregion Remove Lines

        #############    Define inputs for combo boxes     #################
        $OfficeList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\Office.csv"
        $SecurityGroupsList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\SecurityGroups.csv"
        #region removeLine
        #$LogOnScriptList = Import-csv "\\saufs01\it\Enterprise Team\UserManagement\Lists\LogOnScript.csv"
        #endregion removeLine
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
        #region removeLine
        <#
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
        #>
        #endregion removeLine
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
        #region removeLines
        #$NewUser2LogonSelect = New-Object “System.Windows.Forms.Label”;
        #$NewUser2LogonSelect.Location = '20,800'
        #$NewUser2LogonSelect.size = '350,40'
        #endregion removeLines
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
        #region Replace Line
        #$NewUserBox2.Controls.AddRange(@($NewUser2textLabel1,$NewUser2textLabel2,$NewUser2textLabel3,$NewUser2textLabel4,$NewUser2textLabel5,$NewUser2comboBox1,$NewUser2comboBox2,$NewUser2comboBox3,$NewUser2comboBox4,$NewUser2comboBox5,$NewUser2OfficeSelect,$NewUser2DescriptionSelect,$NewUser2DistributionSelect,$NewUser2SecuritySelect,$NewUser2LogonSelect))
        $NewUserBox2.Controls.AddRange(@($NewUser2textLabel1,$NewUser2textLabel2,$NewUser2textLabel3,$NewUser2textLabel4,                     $NewUser2comboBox1,$NewUser2comboBox2,$NewUser2comboBox3,$NewUser2comboBox4,$NewUser2comboBox5,$NewUser2OfficeSelect,$NewUser2DescriptionSelect,$NewUser2DistributionSelect,$NewUser2SecuritySelect))
        #endregion Remplace Line
        $NewUserBox3.Controls.AddRange(@($NewUsertextLabel5,$NewUsertextLabel6,$NewUsertextLabel7,$NewUsertextLabel8,$NewUsertextLabel9,$NewUsertextLabel10))
        #### Activate the form ###
        $NewUserForm1.Add_Shown({$NewUserForm1.Activate()})    
        $dialogResult = $NewUserForm1.ShowDialog()
      
        ########                    set variables               ############# 
        $FirstName = $NewUsertextBox1.text
        $LastName = $NewUsertextBox2.text
        $LogOnName = $NewUsertextBox3.text
        $AccountExpires = $NewUsertextBox4.text
        $DisplayName = $LastName + ", " + $FirstName
        
        ########   Don't accept null username or mailbox     ################ 
        if ($NewUsertextBox1.text -eq "") {
            [System.Windows.Forms.MessageBox]::Show("You need to type in details !!!!!`n`nTrying to enter blank fields is never a good idea.", "Add New Shared Mailbox", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $NewUserForm1.Close()
            $NewUserForm1.Dispose()
            break
        }
      
        ##########  Check to see if Samaccountname is already in use  ###########
        $User = Get-ADUser -Filter {sAMAccountName -eq $LogOnName}
        If ($Null -ne $User) {
            Add-Type -AssemblyName System.Windows.Forms 
            [System.Windows.Forms.MessageBox]::Show("The LogOnName $LogOnName can't be used because it's assigned to an existing user account.`n`nThe next page will display the current usernames in use for $LogOnName`n`nPlease use a LogOnName that's not currently in use.", "ERROR - CAN'T ADD NEW USER",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Get-AdUser -Filter "SamAccountName -like '$LogOnName*'" | Select-Object SamAccountName | Out-GridView -title "User accounts currently in use"
            $NewUserForm1.Close()
            $NewUserForm1.Dispose()
            Remove-Variable DisplayName
            AddNewUser
        } 
        Else {
      
            ##   CHECK - continue if only 1 EmailName in pipe if not exit      ##
            if (($LogOnName | Measure-Object).count -ne 1) {AddNewUser}

            $Password = "Helpdesk123"
            $Office =  $NewUser2OfficeSelect.text
            $Description = $NewUser2DescriptionSelect.Text
            $DistributionGroup = $NewUser2DistributionSelect.Text
            $SecurityGroup = $NewUser2SecuritySelect.Text
            #region RemoveLine
            #$LogOnScript = $NewUser2LogonSelect.Text
            #endregion removeLine
            
            #############        Create AD account         ###########################
            New-AdUser  -Name "$DisplayName" -SamAccountName $LogonName -Path "OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local" –AccountPassword ($Password | 
                ConvertTo-SecureString -AsPlainText –Force)

            #############  Create form to pause for 5 sec   ##########################
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
            $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait. .............."
            $objForm.Controls.Add($objLabel)
            # Show the form
            $objForm.Show()| Out-Null
            # wait 10 seconds
            Start-Sleep -Seconds 5
            # destroy form
            $objForm.Close() | Out-Null
        
            #######    Add security & distribution group permissions    ############## 
            Add-ADGroupMember -PassThru "DomainShareAccess" $LogOnName
            Add-ADGroupMember -PassThru "CN=$Securitygroup,OU=Security Groups,OU=SCS Users,DC=scotcourts,DC=local" $LogOnName 
            #region ChangeLine
            #Add-ADGroupMember -PassThru "$DistributionGroup" $LogOnName
            Add-ADGroupMember -PassThru "CN=$DistributionGroup,ou=distribution lists,ou=groups,ou=courts,dc=scotcourts,dc=local" $LogOnName
            #endregion ChangeLine
            Add-ADGroupMember -PassThru "GPO SF - Folder Redirection 2" $LogOnName

            ###################     Set user AD properties             ############### 
            Set-AdUser –PassThru -Identity $LogonName –GivenName "$FirstName" –Surname "$LastName" -DisplayName "$Displayname" 
            Set-ADUser –PassThru -Identity $LogonName -Office $Office –Description $Description 
            Set-AdUser –PassThru -Identity $LogonName -UserPrincipalName "$LogonName@scotcourts.local"
            #region removeLines
            #Set-ADUser –PassThru -Identity $LogonName –HomeDrive "P:" -HomeDirectory "\\scotcourts.local\data\users\$LogOnName"
            #Set-ADUser –PassThru -Identity $LogonName -ProfilePath "\\scotcourts.local\data\profiles\$LogOnName"
            #endregion
            Set-ADUser –PassThru -Identity $LogonName -AccountExpirationDate $AccountExpires
            #region removeLIne
            #Set-ADUser –PassThru -Identity $LogonName -ScriptPath $LogOnScript
            #endregion removeLine
            ######## Set password change at next logon      ##########################
            #Set-ADUser -Identity $LogonName -ChangePasswordAtLogon $true

            #############       Disable New User account    ##########################
            Set-ADUser -Identity $LogonName  -Enabled $False 

            #############  Create form to pause for 5 sec  ##########################
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
            
            ###########        Create mailbox for user          ######################
            Enable-MailBox -Identity $LogonName@scotcourts.local

            #######   Disable Pop, OWA, Imap & ActiveSync for user ###################
            Set-CASMailbox -Identity $LogonName -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False

            ###########        Generate Form complete           ######################
            Add-Type -AssemblyName System.Windows.Forms 
            $StartMessage = [System.Windows.Forms.MessageBox]::Show("The User account and mailbox have been created in the SCTS\User Accounts\SCTS Users OU.`n`nNote1:  The user account needs to be moved from the SCTS\User Accounts\SCTS Users OU to the correct user OU in AD.:`n`nNote2:  The user account needs to be enabled before use.", "New User Account", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            ###########      Send email to helpdesk             ######################
            Send-MailMessage -To helpdesk@scotcourts.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New User Account $LogOnName added. The new User needs moved out of the New User OU in AD." -Body "A new user account has been added:`n`nUserName:   $DisplayName`n`nLogOnName:   $LogOnName`n`nLocation:   $Description`n`nDistribution List:  $DistributionGroup`n`nSecurity Group:   $SecurityGroup " -SmtpServer mail.scotcourts.local
            Remove-Variable DisplayName
            AddNewUser
        }
    }
}
AddNewUser
