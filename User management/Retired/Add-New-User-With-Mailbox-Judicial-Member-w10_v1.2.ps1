# Add New Judicial User Account with Exchange 2013 mailbox:
# Author         John Mckay
# Date           22/03/2018
# Version        1.2
# Purpose        To add a new judicial user with an exchange 2013 mailbox.
# Usage          Helpdesk staff run a shorcut to Add-New-Judicial-User
# Changes        v1.2  GK This version uses the new unified OU for Windows 10 users as a searchbase and location for new user accounts. 
#                         User intefaces no longer contain fields for roaming profile path, homepath or login script.
#                         Commenting and indenting is updated for readability.
#                         update all unc paths to use dfs form
#                v1.01 GK changes relating to the new personal folder location
#                v1.00 JM Original
 
#######         Create Session with Exchange 2013         ##############
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sauex01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$version='1.2'

############     Set icon for all forms and subforms      ###############
$Icon = "\\scotcourts.local\data\it\Enterprise Team\Usermanagement\icons\user.ico"

#######              Show Start Message:                  ###############
Add-Type -AssemblyName System.Windows.Forms 
[System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”) | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("This script creates a New Judicial User Account with a mailbox in Exchange 2013.`n`nThe New Judicial Account will be created in the SCTS\User Accounts\SCTS Users OU in AD & the account will be disabled.`n`nBefore use the New Judicial User Account needs to be moved from the SCTS\User Accounts\SCTS Users OU to the correct User OU in AD & enabled.`n`nPlease click OK to continue or Cancel to exit", "Add New Judicial User Account with Exchange 2013 mailbox.", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
if ($StartMessage -eq 'Cancel') { exit } 
else {

####     Create SubForm  Add New Shared mailbox Sub Form          ####
Function AddNewJudicialUser{
    #############    Define inputs for combo boxes     #################
    $JudFirstName = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\FirstNameJud.csv"
    $SecurityGroupsList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\SecurityGroups.csv"
    $LogOnScriptList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\LogOnScript.csv"
    $DescriptionList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\descriptionJud.csv"
    $OfficeList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\officeJud.csv"
    $distributionList = Import-csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\distributionJud.csv"
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
    $NewUserForm1.Text = "Add New Judicial User Account with Exchange 2013 mailbox (W10) v$Version"
    $NewUserForm1.Font = New-Object System.Drawing.Font("Ariel",10)
    ### Create group 1 box in form. ####
    $NewUserBox1 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox1.Location = '40,20'
    $NewUserBox1.size = '700,190'
    $NewUserBox1.text = "1. Enter the new Judicial Users details:"
    ### Create group 1 box text labels. ###
    $NewUsertextLabel1 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel1.Location = '20,25'
    $NewUsertextLabel1.size = '350,40'
    $NewUsertextLabel1.Text = "Title  :" 
    $NewUsertextLabel2 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel2.Location = '20,65'
    $NewUsertextLabel2.size = '350,40'
    $NewUsertextLabel2.Text = "Last Name :  (e.g Bloggs with capital 'B')." 
    $NewUsertextLabel3 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel3.Location = '20,102'
    $NewUsertextLabel3.size = '420,40'
    $NewUsertextLabel3.Text = "Initial : (MANDATORY for all Sheriffs NOT REQUIRED for Judges)." 
    $NewUsertextLabel4 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel4.Location = '20,140'
    $NewUsertextLabel4.size = '370,40'
    $NewUsertextLabel4.Text = "Account Expiry Date : (If temporary e.g 25/03/2018)." 
    ### Create group 1 box text boxes. ###
    ##  First name combo  ##
    $NewUsercomboBox1 = New-Object System.Windows.Forms.ComboBox
    $NewUsercomboBox1.Location = '445,20'
    $NewUsercomboBox1.Size = '230,40'
    $NewUsercomboBox1.AutoCompleteMode = 'Suggest'
    $NewUsercomboBox1.AutoCompleteSource = 'ListItems'
    $NewUsercomboBox1.Sorted = $false;
    $NewUsercomboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUsercomboBox1.SelectedItem = $NewUsercomboBox1.Items[0]
    $NewUsercomboBox1.DataSource = $JudFirstName.FirstName 
    $NewUsercomboBox1.add_SelectedIndexChanged({$NewUsertextLabel12.Text = $NewUsertextLabel9.Text = "$($NewUsercomboBox1.SelectedItem.ToString())"})
    ##  Last name textbox ##
    $NewUsertextBox2 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox2.Location = '445,60'
    $NewUsertextBox2.Size = '230,40'
    $NewUsertextBox2.add_textChanged({$NewUsertextLabel13.Text = $NewUsertextLabel10.Text = "$($NewUsertextBox2.text)"})
    ##   Initial textbox  ##
    $NewUsertextBox3 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox3.Location = '445,95'
    $NewUsertextBox3.Size = '230,40'
    $NewUsertextBox3.add_TextChanged({$NewUsertextLabel14.Text = $NewUsertextLabel11.Text = "$($NewUsertextBox3.text)"})
    ## Acc Expire textbox ##
    $NewUsertextBox4 = New-Object System.Windows.Forms.TextBox
    $NewUsertextBox4.Location = '445,135'
    $NewUsertextBox4.Size = '230,40'
    $NewUsertextBox4.add_textChanged({$AccountExpires = "$($NewUsertextBox4.text)"})
    ### Create group 2 box in form. ###
    $NewUserBox2 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox2.Location = '40,230'
    $NewUserBox2.size = '700,280'
    $NewUserBox2.text = "2. Enter the new Judicial users AD details:"
    ### Create group 2 box text labels. ###
    $NewUser2textLabel1 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel1.Location = '20,35'
    $NewUser2textLabel1.size = '350,40'
    $NewUser2textLabel1.Text = "Office :" 
    $NewUser2textLabel2 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel2.Location = '20,75'
    $NewUser2textLabel2.size = '350,40'
    $NewUser2textLabel2.Text = "Description :" 
    $NewUser2textLabel3 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel3.Location = '20,115'
    $NewUser2textLabel3.size = '350,40'
    $NewUser2textLabel3.Text = "Distribution List 1 :" 
    $NewUser2textLabel4 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel4.Location = '20,195'
    $NewUser2textLabel4.size = '350,40'
    $NewUser2textLabel4.Text = "Security Group (to access p and s drives):" 
    $NewUser2textLabel5 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel5.Location = '20,235'
    $NewUser2textLabel5.size = '250,40'
    $NewUser2textLabel5.Text = "LogOn Script :"
    $NewUser2textLabel6 = New-Object System.Windows.Forms.Label;
    $NewUser2textLabel6.Location = '20,155'
    $NewUser2textLabel6.size = '350,40'
    $NewUser2textLabel6.Text = "Distribution List 2 :" 
    ### Create group 2 box combo boxes. ###
    ###  Populate "Office" ComboBox1   ###
    $NewUser2comboBox1 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox1.Location = '375,30'
    $NewUser2comboBox1.Size = '300,40'
    $NewUser2comboBox1.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox1.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox1.Sorted = $false;
    $NewUser2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox1.SelectedItem = $NewUser2comboBox1.Items[0]
    $NewUser2comboBox1.DataSource = $OfficeList.Office 
    $NewUser2comboBox1.add_SelectedIndexChanged({$NewUser2OfficeSelect.Text = "$($NewUser2comboBox1.SelectedItem.ToString())"})
    ###  Populate "Description" ComboBox2 ###
    $NewUser2comboBox2 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox2.Location = '375,70'
    $NewUser2comboBox2.Size = '300,40'
    $NewUser2comboBox2.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox2.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox2.Sorted = $false;
    $NewUser2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox2.SelectedItem = $NewUser2comboBox2.Items[0]
    $NewUser2comboBox2.DataSource = $DescriptionList.Description 
    $NewUser2comboBox2.add_SelectedIndexChanged({$NewUser2DescriptionSelect.Text = "$($NewUser2comboBox2.SelectedItem.ToString())"})
    ###  Populate "Distribution List" ComboBox3 ###
    $NewUser2comboBox3 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox3.Location = '375,110'
    $NewUser2comboBox3.Size = '300,40'
    $NewUser2comboBox3.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox3.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox3.SelectedItem = $NewUser2comboBox3.Items[0]
    $NewUser2comboBox3.DataSource = $distributionList.distribution  
    $NewUser2comboBox3.add_SelectedIndexChanged({$NewUser2DistributionSelect.Text = "$($NewUser2comboBox3.SelectedItem.ToString())"})
    ###  Populate "Security Group" ComboBox4 ###
    $NewUser2comboBox4 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox4.Location = '375,190'
    $NewUser2comboBox4.Size = '300,40'
    $NewUser2comboBox4.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox4.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox4.Sorted = $false;
    $NewUser2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox4.SelectedItem = $NewUser2comboBox4.Items[0]
    $NewUser2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
    $NewUser2comboBox4.add_SelectedIndexChanged({$NewUser2SecuritySelect.Text = "$($NewUser2comboBox4.SelectedItem.ToString())"})
    ###  Populate "LogOn Script" ComboBox5 ###
    # hide this portion of the GUI
    <#
    $NewUser2comboBox5 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox5.Location = '375,230'
    $NewUser2comboBox5.Size = '300,40'
    $NewUser2comboBox5.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox5.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox5.Sorted = $false;
    $NewUser2comboBox5.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox5.SelectedItem = $NewUser2comboBox5.Items[0]
    $NewUser2comboBox5.DataSource = $LogOnScriptList.Logonscript 
    $NewUser2comboBox5.add_SelectedIndexChanged({$NewUser2LogonSelect.Text = "$($NewUser2comboBox5.SelectedItem.ToString())"})
    #>
    ###  Populate "Description2" ComboBox2 ###
    $NewUser2comboBox6 = New-Object System.Windows.Forms.ComboBox
    $NewUser2comboBox6.Location = '375,150'
    $NewUser2comboBox6.Size = '300,40'
    $NewUser2comboBox6.AutoCompleteMode = 'Suggest'
    $NewUser2comboBox6.AutoCompleteSource = 'ListItems'
    $NewUser2comboBox6.Sorted = $false;
    $NewUser2comboBox6.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $NewUser2comboBox6.SelectedItem = $NewUser2comboBox6.Items[0]
    $NewUser2comboBox6.DataSource = $DistributionList.Distribution 
    $NewUser2comboBox6.add_SelectedIndexChanged({$NewUser2Distribution2Select.Text = "$($NewUser2comboBox6.SelectedItem.ToString())"})
    ### Create group 2 labels to take combobox output ###
    $NewUser2OfficeSelect = New-Object System.Windows.Forms.Label;
    $NewUser2OfficeSelect.Location = '20,600'
    $NewUser2OfficeSelect.size = '350,40'
    $NewUser2DescriptionSelect = New-Object System.Windows.Forms.Label;
    $NewUser2DescriptionSelect.Location = '20,650'
    $NewUser2DescriptionSelect.size = '350,40'
    $NewUser2DistributionSelect = New-Object System.Windows.Forms.Label;
    $NewUser2DistributionSelect.Location = '20,700'
    $NewUser2DistributionSelect.size = '350,40'
    $NewUser2SecuritySelect = New-Object System.Windows.Forms.Label;
    $NewUser2SecuritySelect.Location = '20,750'
    $NewUser2SecuritySelect.size = '350,40'
    # remove this Block
    <#
    $NewUser2LogonSelect = New-Object System.Windows.Forms.Label;
    $NewUser2LogonSelect.Location = '20,800'
    $NewUser2LogonSelect.size = '350,40'
    #>
    $NewUser2Distribution2Select = New-Object System.Windows.Forms.Label;
    $NewUser2Distribution2Select.Location = '20,800'
    $NewUser2Distribution2Select.size = '350,40'
    ### Create group 3 box in form. ###
    $NewUserBox3 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox3.Location = '40,535'
    $NewUserBox3.size = '700,95'
    $NewUserBox3.text = "3. Check the details below are correct before proceeding:"
    ### Create group 3 box text labels.
    ## message label ##
    $NewUsertextLabel5 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel5.Location = '20,30'
    $NewUsertextLabel5.size = '350,30'
    $NewUsertextLabel5.Text = "New User will appear in AD and Global Address List as:" 
    ## DisplayName label ##
    $NewUsertextLabel6 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel6.Location = '20,60'
    $NewUsertextLabel6.Size = '300,30'
    $NewUsertextLabel6.ForeColor = "Blue"
    ## message label ##
    $NewUsertextLabel7 = New-Object System.Windows.Forms.Label;
    $NewUsertextLabel7.Location = '430,30'
    $NewUsertextLabel7.size = '200,30'
    $NewUsertextLabel7.Text = "With the LogOnName:"
    ## LogOnName label ##
    $NewUsertextLabel8 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel8.Location = '460,60'
    $NewUsertextLabel8.Size = '200,30'
    $NewUsertextLabel8.ForeColor = "Blue"
    ## First name label ##
    $NewUsertextLabel9 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel9.Location = '20,300'
    $NewUsertextLabel9.Size = '400,30'
    $NewUsertextLabel9.ForeColor = "Blue"
    $NewUsertextLabel9.add_TextChanged({$NewUsertextLabel6.Text = "$($NewUsertextLabel9.text + " " + $NewUsertextLabel10.text + " " + $NewUsertextLabel11.text)"})
    ## Last name label ##
    $NewUsertextLabel10 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel10.Location = '20,350'
    $NewUsertextLabel10.Size = '250,30'
    $NewUsertextLabel10.ForeColor = "Blue"
    $NewUsertextLabel10.add_TextChanged({$NewUsertextLabel6.Text = "$($NewUsertextLabel9.text + " " + $NewUsertextLabel10.text + " " + $NewUsertextLabel11.text)"})
    ##  Initial label  ##
    $NewUsertextLabel11 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel11.Location = '20,375'
    $NewUsertextLabel11.Size = '250,30'
    $NewUsertextLabel11.ForeColor = "Blue"
    $NewUsertextLabel11.add_TextChanged({$NewUsertextLabel6.Text = "$($NewUsertextLabel9.text + " " + $NewUsertextLabel10.text + " " + $NewUsertextLabel11.text)"})
    ## First name label ##
    $NewUsertextLabel12 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel12.Location = '20,400'
    $NewUsertextLabel12.Size = '250,50'
    $NewUsertextLabel12.ForeColor = "Blue"
    $NewUsertextLabel12.add_TextChanged({$NewUsertextLabel8.Text = "$($NewUsertextLabel9.text + $NewUsertextLabel11.text + $NewUsertextLabel10.text)"})
    ## Last name label ##
    $NewUsertextLabel13 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel13.Location = '20,425'
    $NewUsertextLabel13.Size = '250,50'
    $NewUsertextLabel13.ForeColor = "Blue"
    $NewUsertextLabel13.add_TextChanged({$NewUsertextLabel8.Text = "$($NewUsertextLabel9.text + $NewUsertextLabel11.text + $NewUsertextLabel10.text)"})
    ##  Initial label  ##
    $NewUsertextLabel14 = New-Object System.Windows.Forms.Label
    $NewUsertextLabel14.Location = '20,450'
    $NewUsertextLabel14.Size = '250,50'
    $NewUsertextLabel14.ForeColor = "Blue"
    $NewUsertextLabel14.add_TextChanged({$NewUsertextLabel8.Text = "$($NewUsertextLabel9.text + $NewUsertextLabel11.text + $NewUsertextLabel10.text)"})
    ### Create group 4 box in form. ###
    $NewUserBox4 = New-Object System.Windows.Forms.GroupBox
    $NewUserBox4.Location = '40,640'
    $NewUserBox4.size = '700,30'
    $NewUserBox4.text = "4. Click Continue or Exit:"
    $NewUserBox4.button
    ### Add an OK button ###
    $ContinueButton = new-object System.Windows.Forms.Button
    $ContinueButton.Location = '640,680'
    $ContinueButton.Size = '100,40'          
    $ContinueButton.Text = 'Continue'
    $ContinueButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
    ### Add a cancel button ###
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '525,680'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Exit"
    $CancelButton.add_Click({
    #$NewUserForm1.Close()
    #$NewUserForm1.Dispose()
    $CancelButton.[System.Environment]::Exit(0)})
    ### Add all the Form controls ### 
    $NewUserForm1.Controls.AddRange(@($NewUserBox1,$NewUserBox2,$NewUserBox3,$NewUserBox4,$ContinueButton,$CancelButton))
    #### Add all the GroupBox controls ###
    $NewUserBox1.Controls.AddRange(@($NewUsertextLabel1,$NewUsertextLabel2,$NewUsertextLabel3,$NewUsertextLabel4,$NewUserComboBox1,$NewUsertextBox2,$NewUsertextBox3,$NewUsertextBox4))
    # replace this Line to remove $NewUser2LogonSelect
    # $NewUserBox2.Controls.AddRange(@($NewUser2textLabel1,$NewUser2textLabel2,$NewUser2textLabel3,$NewUser2textLabel4,$NewUser2textLabel5,$NewUser2textLabel6,$NewUser2comboBox1,$NewUser2comboBox2,$NewUser2comboBox3,$NewUser2comboBox4,$NewUser2comboBox5,$NewUser2comboBox6,$NewUser2OfficeSelect,$NewUser2DescriptionSelect,$NewUser2DistributionSelect,$NewUser2SecuritySelect,$NewUser2LogonSelect,$NewUser2Distribution2Select))
    $NewUserBox2.Controls.AddRange(@($NewUser2textLabel1,$NewUser2textLabel2,$NewUser2textLabel3,$NewUser2textLabel4,$NewUser2textLabel5,$NewUser2textLabel6,$NewUser2comboBox1,$NewUser2comboBox2,$NewUser2comboBox3,$NewUser2comboBox4,$NewUser2comboBox5,$NewUser2comboBox6,$NewUser2OfficeSelect,$NewUser2DescriptionSelect,$NewUser2DistributionSelect,$NewUser2SecuritySelect,$NewUser2Distribution2Select))
    $NewUserBox3.Controls.AddRange(@($NewUsertextLabel5,$NewUsertextLabel6,$NewUsertextLabel7,$NewUsertextLabel8,$NewUsertextLabel9,$NewUsertextLabel10,$NewUsertextLabel11,$NewUsertextLabel12,$NewUsertextLabel13,$NewUsertextLabel14))
    #### Activate the form ###
    $NewUserForm1.Add_Shown({$NewUserForm1.Activate()})    
    $dialogResult = $NewUserForm1.ShowDialog()
 
    ########                    set variables               ############# 
    $FirstName = $NewUserComboBox1.text
    $LastName = $NewUsertextBox2.text
    $Initial = $NewUsertextBox3.text
    $AccountExpires = $NewUsertextBox4.text

    ########   Don't accept null username or mailbox     ################ 
    if ($NewUsertextBox2.text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("You need to type in details !!!!!`n`nTrying to enter blank fields is never a good idea.", "Add New Judicial User", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        Remove-Variable FirstName, LastName, Initial
        $NewUserForm1.Close()
        $NewUserForm1.Dispose()
        break
    }
    ########             check sheriff has initial          ############# 
    ElseIf (($Firstname -eq "sheriff") -or ($Firstname -eq "sheriffs") -or ($Firstname -eq "sheriffp")  -and ($Initial -eq "")){
        [System.Windows.Forms.MessageBox]::Show("You need to type in an INITIAL !!!!!`n`nSheriffs need to have an initial.", "Add New Judicial User", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Remove-Variable FirstName, LastName, Initial
        $NewUserForm1.Close()
        $NewUserForm1.Dispose()
        AddNewJudicialUser
    }
     
    ########          check judges have no initial          ############# 
    If (($FirstName -eq "Lord") -or ($FirstName -eq "Lady") -and ($Initial -ne "")){
        [System.Windows.Forms.MessageBox]::Show("You have selected a Judge with an INITIAL !!!!!`n`nJudges don't have initials added to their accounts.", 
            "Add New Judicial User", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Remove-Variable FirstName, LastName, Initial
        $NewUserForm1.Close()
        $NewUserForm1.Dispose()
        AddNewJudicialUser
    }
     
    ########      remove space from displayname for Judges  ############# 
    If (($FirstName -eq "Lord") -or ($FirstName -eq "Lady")){
        $DisplayName = $FirstName + " " + $LastName
        $LogOnName = $FirstName + $LastName}
    Else {
        $DisplayName = $NewUsertextLabel6.Text
        $LogOnName = $NewUsertextLabel8.Text
    }

    ##########  Check to see if Samaccountname is already in use  ###########
    $User = Get-ADUser -Filter {sAMAccountName -eq $LogOnName}
    If ($Null -ne $User  ) {
        Add-Type -AssemblyName System.Windows.Forms 
        [System.Windows.Forms.MessageBox]::Show("The LogOnName $LogOnName can't be used because it's assigned to an existing judicial user account.`n`nThe next page will display the current usernames in use for $LogOnName`n`nPlease use a LogOnName that's not currently in use`n`nCheck if user would like another initial added or use first name as initial.", "ERROR - CAN'T ADD NEW JUDICIAL USER", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Get-AdUser -Filter "SamAccountName -like '$LogOnName*'" | Select-Object SamAccountName | Out-GridView -title "User accounts currently in use"
        #Remove-Variable FirstName, LastName, Initial, DisplayName
        $NewUserForm1.Close()
        $NewUserForm1.Dispose()
        AddNewJudicialUser}
    Else {    
        ##   CHECK - continue if only 1 EmailName in pipe if not exit      ##
        if (($LogOnName | Measure-Object).count -ne 1) {AddNewJudicialUser}
        
        $Password = "Helpdesk123"
        $Office =  $NewUser2OfficeSelect.text
        $Description = $NewUser2DescriptionSelect.Text
        $DistributionGroup = $NewUser2DistributionSelect.Text
        $DistributionGroup2 = $NewUser2Distribution2Select.Text
        $SecurityGroup = $NewUser2SecuritySelect.Text
        # remove this line - no more logon scripts
        #$LogOnScript = $NewUser2LogonSelect.Text
    }
    
    #############        Create AD account         ###########################
    # change OU path for new accounts - Windows 10 ready.
    New-AdUser  -Name "$DisplayName" -SamAccountName $LogonName -Path "OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local" –AccountPassword ($Password | ConvertTo-SecureString -AsPlainText –Force)
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
    $objLabel.Text = "A New Judicial User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS\User Accounts\SCTS Users OU.`n`nPlease Wait. .............."
    $objForm.Controls.Add($objLabel)
    # Show the form
    $objForm.Show()| Out-Null
    # wait 10 seconds
    Start-Sleep -Seconds 5
    # close form
    $objForm.Close() | Out-Null
    
    #######    Add security & distribution group permissions    ############## 
    Add-ADGroupMember -PassThru "DomainShareAccess" $LogOnName
    Add-ADGroupMember -PassThru "All Judicial Studies Access" $LogOnName
    Add-ADGroupMember -PassThru "Judicial Studies Access" $LogOnName
    Add-ADGroupMember -PassThru "CN=$Securitygroup,OU=Security Groups,OU=SCS Users,DC=scotcourts,DC=local" $LogOnName 
    Add-ADGroupMember -PassThru "$DistributionGroup" $LogOnName
    Add-ADGroupMember -PassThru "$DistributionGroup2" $LogOnName
    Add-ADGroupMember -PassThru "GPO SF - Folder Redirection 2" $LogOnName
     
    ###################     Set user AD properties             ############### 
    Set-AdUser –PassThru -Identity $LogonName –GivenName "$FirstName" –Surname "$LastName" -DisplayName "$Displayname" 
    Set-ADUser –PassThru -Identity $LogonName -Office $Office –Description $Description -Initials $Initial 
    Set-AdUser –PassThru -Identity $LogonName -UserPrincipalName "$LogonName@scotcourts.local"
    # no more profile path
    #Set-ADUser –PassThru -Identity $LogonName -ProfilePath "\\scotcourts.local\data\profiles\$LogOnName"
    Set-ADUser –PassThru -Identity $LogonName -AccountExpirationDate $AccountExpires
    # no more logon script
    #Set-ADUser –PassThru -Identity $LogonName -ScriptPath $LogOnScript
     
    ######## Set password change at next logon      ##########################
    #Set-ADUser -Identity $LogonName -ChangePasswordAtLogon $true
    
    #############       Disable New User account    ##########################
    Set-ADUser -Identity $LogonName  -Enabled $False 
    ##########################################################################
    #############  Create form to pause for 5 sec  ##########################
    ##########################################################################
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
    $objLabel.Text = "A New Judicial User Mailbox is being created in Exchange 2013 with the details you entered. `n`n`n`nPlease Wait. .............."
    $objForm.Controls.Add($objLabel)
    # Show the form
    $objForm.Show()| Out-Null
    # wait 5 seconds
    Start-Sleep -Seconds 5
    # close form
    $objForm.Close() | Out-Null

    #Remove This Block - P drive all this is now done by GPO on first login.
    ##########################################################################
    ###########     Add users P drive folder on saufs01    ###################
    ##########################################################################
    #New-Item -Path \\saufs01\users -Name $LogonName -ItemType Directory -Force
    ##########################################################################
    #####   Add user as owner on p drive & set to inherit permissions    #####
    ##########################################################################
    #$acl = Get-Acl \\saufs01\users\$LogOnName
    #$acl.SetAccessRuleProtection($false, $false)
    #$acl.SetOwner([System.Security.Principal.NTAccount]"$LogOnName")
    #Set-Acl \\saufs01\users\$LogOnName $acl
    ##########################################################################
    ###########         Set permissions complete           ###################
    ##########################################################################
    #
     
    ###########        Create mailbox for user          ######################
    Enable-MailBox -Identity $LogonName@scotcourts.local


    #######   Disable Pop, OWA, Imap & ActiveSync for user ###################
    Set-CASMailbox -Identity $LogonName -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
    
    ###########        Generate Form complete           ######################
    Add-Type -AssemblyName System.Windows.Forms 
    [System.Windows.Forms.MessageBox]::Show("The Judicial User account and mailbox have been created in the 'SCTS\User Accounts\SCTS Users' OU.`n`nNote1:  The Judicial user account needs to be enabled before use.", "New Judicial User Account", 
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    ###########      Send email to helpdesk             ######################
    Send-MailMessage -To helpdesk@scotcourts.gov.uk -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: New Judicial User Account $LogOnName added. The new Judicial User needs moved out of the New User OU in AD." -Body "A new Judicial user account has been added:`n`nUserName:   $DisplayName`n`nLogOnName:   $LogOnName`n`nLocation:   $Description`n`nDistribution List:  $DistributionGroup`n$DistributionGroup2`n`nSecurity Group:   $SecurityGroup " -SmtpServer mail.scotcourts.local
    Remove-Variable FirstName, LastName, Initial, DisplayName
    $NewUserForm1.Close()
    $NewUserForm1.Dispose()
    AddNewJudicialUser
}
}  
    AddNewJudicialUser
    