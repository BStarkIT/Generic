#
# Set icon for all forms and subforms
#
$Icon = '\\saufs01\IT\Enterprise Team\Usermanagement\icons\User.ico'
#
# Get listof UserNames from AD
#
$UserNameList = Get-ADUser –filter * -searchbase 'ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
#
#
function AdminType {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    ### Set the details of the form. ###
    $AdminSelectForm = New-Object System.Windows.Forms.Form
    $AdminSelectForm.width = 730
    $AdminSelectForm.height = 550
    $AdminSelectForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $AdminSelectForm.Controlbox = $false
    $AdminSelectForm.Icon = $Icon
    $AdminSelectForm.FormBorderStyle = 'Fixed3D'
    $AdminSelectForm.Text = 'Set Type/team.'
    $AdminSelectForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    #
    $AdminSelectFormBox1 = New-Object System.Windows.Forms.GroupBox
    $AdminSelectFormBox1.Location = '10,10'
    $AdminSelectFormBox1.size = '340,170'
    $AdminSelectFormBox1.text = 'Current settings:'
    $AdminSelectFormtextLabel1 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtextLabel1.Location = '20,20'
    $AdminSelectFormtextLabel1.size = '100,20'
    $AdminSelectFormtextLabel1.Text = 'UserName:' 
    $AdminSelectFormtextLabel2 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtextLabel2.Location = '20,50'
    $AdminSelectFormtextLabel2.size = '100,20'
    $AdminSelectFormtextLabel2.Text = 'Display Name:'
    $AdminSelectFormtextLabel3 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtextLabel3.Location = '20,80'
    $AdminSelectFormtextLabel3.size = '100,20'
    $AdminSelectFormtextLabel3.Text = 'First Name:'  
    $AdminSelectFormtextLabel4 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtextLabel4.Location = '20,110'
    $AdminSelectFormtextLabel4.size = '100,20'
    $AdminSelectFormtextLabel4.Text = 'Last Name:'   
    $AdminSelectFormtext1 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtext1.Location = '120,20'
    $AdminSelectFormtext1.size = '150,20'
    $AdminSelectFormtext1.Text = $UserSamAccountName
    $AdminSelectFormtext1.ForeColor = 'Blue'
    $AdminSelectFormtext2 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtext2.Location = '120,50'
    $AdminSelectFormtext2.size = '200,20'
    $AdminSelectFormtext2.Text = $SelectedUser
    $AdminSelectFormtext2.ForeColor = 'Blue'
    $AdminSelectFormtext3 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtext3.Location = '120,80'
    $AdminSelectFormtext3.size = '150,20'
    $AdminSelectFormtext3.Text = $UserFirstName  
    $AdminSelectFormtext3.ForeColor = 'Blue'
    $AdminSelectFormtext4 = New-Object System.Windows.Forms.Label
    $AdminSelectFormtext4.Location = '120,110'
    $AdminSelectFormtext4.size = '200,20'
    $AdminSelectFormtext4.Text = $UserSurName 
    $AdminSelectFormtext4.ForeColor = 'Blue'
    #
    $AdminSelectFormBox2 = New-Object System.Windows.Forms.GroupBox
    $AdminSelectFormBox2.Location = '360,10'
    $AdminSelectFormBox2.size = '340,170'
    $AdminSelectFormBox2.text = 'New Settings:'
    $AdminSelectFormB2textLabel1 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2textLabel1.size = '100,20'
    $AdminSelectFormB2textLabel1.Location = '20,20'
    $AdminSelectFormB2textLabel1.Text = 'UserName:' 
    $AdminSelectFormB2textLabel2 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2textLabel2.Location = '20,50'
    $AdminSelectFormB2textLabel2.size = '100,20'
    $AdminSelectFormB2textLabel2.Text = 'Display Name:'
    $AdminSelectFormB2textLabel3 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2textLabel3.Location = '20,80'
    $AdminSelectFormB2textLabel3.size = '100,20'
    $AdminSelectFormB2textLabel3.Text = 'Password:'  
    $AdminSelectFormB2text1 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2text1.Location = '120,20'
    $AdminSelectFormB2text1.size = '150,20'
    $AdminSelectFormB2text1.ForeColor = 'Green'
    $AdminSelectFormB2text1.Text = $UserSamAccountNameadmin
    $AdminSelectFormB2text2 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2text2.Location = '120,50'
    $AdminSelectFormB2text2.size = '200,20'
    $AdminSelectFormB2text2.Text = $SelectedAdminUser
    $AdminSelectFormB2text2.ForeColor = 'Green'
    $AdminSelectFormB2text3 = New-Object System.Windows.Forms.Label
    $AdminSelectFormB2text3.Location = '120,80'
    $AdminSelectFormB2text3.size = '150,20'
    $AdminSelectFormB2text3.Text = "TTyy7788!1"
    $AdminSelectFormB2text3.ForeColor = 'Green' 
    #
    # Menu boxes
    $AdminSelectFormBox3 = New-Object System.Windows.Forms.GroupBox
    $AdminSelectFormBox3.Location = '10,190'
    $AdminSelectFormBox3.size = '690,220'
    $AdminSelectFormBox3.text = 'Select Team:'
    $BoxRadioButton01 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton01.Location = '20,210'
    $BoxRadioButton01.size = '200,30'
    $BoxRadioButton01.Checked = $false 
    $BoxRadioButton01.Text = 'SCTS-Admins.'
    $BoxRadioButton02 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton02.Location = '20,235'
    $BoxRadioButton02.size = '200,30'
    $BoxRadioButton02.Checked = $false
    $BoxRadioButton02.Text = 'SCTS-ElevatedAdmins.'
    $BoxRadioButton03 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton03.Location = '20,260'
    $BoxRadioButton03.size = '200,30'
    $BoxRadioButton03.Checked = $false
    $BoxRadioButton03.Text = 'SCTS-RemoteDesktop.'
    $BoxRadioButton04 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton04.Location = '20,285'
    $BoxRadioButton04.size = '200,30'
    $BoxRadioButton04.Checked = $false 
    $BoxRadioButton04.Text = 'SCTS-ServerAdmins.'
    $BoxRadioButton05 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton05.Location = '20,310'
    $BoxRadioButton05.size = '200,30'
    $BoxRadioButton05.Checked = $false
    $BoxRadioButton05.Text = 'SCTS-Helpdesk.'
    $BoxRadioButton06 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton06.Location = '20,335'
    $BoxRadioButton06.size = '200,30'
    $BoxRadioButton06.Checked = $false
    $BoxRadioButton06.Text = 'SCTS-Field Service.'
    $BoxRadioButton07 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton07.Location = '20,360'
    $BoxRadioButton07.size = '200,30'
    $BoxRadioButton07.Checked = $false 
    $BoxRadioButton07.Text = 'SCTS-UnifiedComms.'
    $BoxRadioButton08 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton08.Location = '220,210'
    $BoxRadioButton08.size = '200,30'
    $BoxRadioButton08.Checked = $false 
    $BoxRadioButton08.Text = 'SCTS-SQLDBA.'
    $BoxRadioButton09 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton09.Location = '220,235'
    $BoxRadioButton09.size = '200,30'
    $BoxRadioButton09.Checked = $false
    $BoxRadioButton09.Text = 'SCTS-HelpdeskManagers.'
    $BoxRadioButton10 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton10.Location = '220,260'
    $BoxRadioButton10.size = '200,30'
    $BoxRadioButton10.Checked = $false
    $BoxRadioButton10.Text = 'SCTS-Architect.'
    $BoxRadioButton11 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton11.Location = '220,285'
    $BoxRadioButton11.size = '200,30'
    $BoxRadioButton11.Checked = $false 
    $BoxRadioButton11.Text = 'SCTS-Audit.'
    $BoxRadioButton12 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton12.Location = '220,310'
    $BoxRadioButton12.size = '200,30'
    $BoxRadioButton12.Checked = $false
    $BoxRadioButton12.Text = 'SCTS-Developer.'
    $BoxRadioButton13 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton13.Location = '220,335'
    $BoxRadioButton13.size = '200,30'
    $BoxRadioButton13.Checked = $false
    $BoxRadioButton13.Text = 'SCTS-DeveloperManager.'
    $BoxRadioButton14 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton14.Location = '220,360'
    $BoxRadioButton14.size = '200,30'
    $BoxRadioButton14.Checked = $false 
    $BoxRadioButton14.Text = 'SCTS-Tester.'
    $BoxRadioButton15 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton15.Location = '420,210'
    $BoxRadioButton15.size = '200,30'
    $BoxRadioButton15.Checked = $false 
    $BoxRadioButton15.Text = 'SCTS-TesterManager.'
    $BoxRadioButton16 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton16.Location = '420,235'
    $BoxRadioButton16.size = '200,30'
    $BoxRadioButton16.Checked = $false
    $BoxRadioButton16.Text = 'SCTS-CyberOperations.'
    $BoxRadioButton17 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton17.Location = '420,260'
    $BoxRadioButton17.size = '240,30'
    $BoxRadioButton17.Checked = $false
    $BoxRadioButton17.Text = 'SCTS-CyberOperationsManager.'
    $BoxRadioButton18 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton18.Location = '420,285'
    $BoxRadioButton18.size = '200,30'
    $BoxRadioButton18.Checked = $false 
    $BoxRadioButton18.Text = 'SCTS-WebDeveloper.'
    $BoxRadioButton19 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton19.Location = '420,310'
    $BoxRadioButton19.size = '200,30'
    $BoxRadioButton19.Checked = $false
    $BoxRadioButton19.Text = 'SCTS-Networks.'
    $BoxRadioButton20 = New-Object System.Windows.Forms.RadioButton
    $BoxRadioButton20.Location = '420,335'
    $BoxRadioButton20.size = '200,30'
    $BoxRadioButton20.Checked = $false
    $BoxRadioButton20.Text = 'SCTS-NetworksManager.'
    $BoxRadioButton01.Add_Click( {
            $BoxRadioButton01.Checked = $true
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton02.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $true
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton03.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $true
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton04.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $true
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton05.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $true
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton06.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $true 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton07.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $true
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton08.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $true
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton09.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $true
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton10.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $true
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton11.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $true
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton12.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $true
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton13.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $true 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton14.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $true
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton15.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $true
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton16.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $true
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton17.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $true
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton18.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $true
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton19.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $true
            $BoxRadioButton20.Checked = $false })
    $BoxRadioButton20.Add_Click( {
            $BoxRadioButton01.Checked = $false
            $BoxRadioButton02.Checked = $false
            $BoxRadioButton03.Checked = $false
            $BoxRadioButton04.Checked = $false
            $BoxRadioButton05.Checked = $false
            $BoxRadioButton06.Checked = $false 
            $BoxRadioButton07.Checked = $false
            $BoxRadioButton08.Checked = $false
            $BoxRadioButton09.Checked = $false
            $BoxRadioButton10.Checked = $false
            $BoxRadioButton11.Checked = $false
            $BoxRadioButton12.Checked = $false
            $BoxRadioButton13.Checked = $false 
            $BoxRadioButton14.Checked = $false
            $BoxRadioButton15.Checked = $false
            $BoxRadioButton16.Checked = $false
            $BoxRadioButton17.Checked = $false
            $BoxRadioButton18.Checked = $false
            $BoxRadioButton19.Checked = $false
            $BoxRadioButton20.Checked = $true })
    $AdminSelectForm.Controls.Add($BoxRadioButton01)
    $AdminSelectForm.Controls.Add($BoxRadioButton02)
    $AdminSelectForm.Controls.Add($BoxRadioButton03)
    $AdminSelectForm.Controls.Add($BoxRadioButton04)
    $AdminSelectForm.Controls.Add($BoxRadioButton05)
    $AdminSelectForm.Controls.Add($BoxRadioButton06)
    $AdminSelectForm.Controls.Add($BoxRadioButton07)
    $AdminSelectForm.Controls.Add($BoxRadioButton08)
    $AdminSelectForm.Controls.Add($BoxRadioButton09)
    $AdminSelectForm.Controls.Add($BoxRadioButton10)
    $AdminSelectForm.Controls.Add($BoxRadioButton11)
    $AdminSelectForm.Controls.Add($BoxRadioButton12)
    $AdminSelectForm.Controls.Add($BoxRadioButton13)
    $AdminSelectForm.Controls.Add($BoxRadioButton14)
    $AdminSelectForm.Controls.Add($BoxRadioButton15)
    $AdminSelectForm.Controls.Add($BoxRadioButton16)
    $AdminSelectForm.Controls.Add($BoxRadioButton17)
    $AdminSelectForm.Controls.Add($BoxRadioButton18)
    $AdminSelectForm.Controls.Add($BoxRadioButton19)
    $AdminSelectForm.Controls.Add($BoxRadioButton20)
    #
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,450'
    $OKButton.Size = '100,40'          
    $OKButton.Text = 'Ok'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '600,450'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $AdminSelectForm.Close()
            $AdminSelectForm.Dispose()
            Return MainForm })
    $TicketBoxlabel = New-Object System.Windows.Forms.Label
    $TicketBoxlabel.Location = New-Object System.Drawing.Point(10, 425)
    $TicketBoxlabel.Size = New-Object System.Drawing.Size(80, 20)
    $TicketBoxlabel.Text = 'Ticket No:'
    $AdminSelectForm.Controls.Add($TicketBoxlabel)
    $TicketBox = New-Object System.Windows.Forms.TextBox
    $TicketBox.Location = New-Object System.Drawing.Point(100, 425)
    $TicketBox.Size = New-Object System.Drawing.Size(80, 20)
    $AdminSelectForm.Controls.Add($TicketBox)
    # Add all the Form controls on one line 
    $AdminSelectForm.Controls.AddRange(@($AdminSelectFormBox1, $AdminSelectFormBox3, $AdminSelectFormBox2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $AdminSelectFormBox1.Controls.AddRange(@($AdminSelectFormtextLabel1, $AdminSelectFormtextLabel2, $AdminSelectFormtextLabel3, $AdminSelectFormtextLabel4, $AdminSelectFormtext1, $AdminSelectFormtext2, $AdminSelectFormtext3, $AdminSelectFormtext4))
    $AdminSelectFormBox2.Controls.AddRange(@($AdminSelectFormB2textLabel1, $AdminSelectFormB2textLabel2, $AdminSelectFormB2textLabel3, $AdminSelectFormB2textLabel5, $AdminSelectFormB2text1, $AdminSelectFormB2text2, $AdminSelectFormB2text3))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $AdminSelectForm.AcceptButton = $OKButton
    $AdminSelectForm.CancelButton = $CancelButton
    # Activate the form
    $AdminSelectForm.Topmost = $true
    $AdminSelectForm.Add_Shown( { $TicketBox.Select() })
    $result = $AdminSelectForm.ShowDialog() 
    # Get the results from the button click
    $ticket = $TicketBox.Text
    if ($Result -eq 'OK') {
        if ($TicketBox.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to enter a Ticket Number!  Trying to enter blank fields is never a good idea.", 'New Admin', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
        Else { 
            if ($BoxRadioButton01.Checked) {
                $Role = "ROLE-SCTS-Admins"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton02.Checked) {
                $Role = "ROLE-SCTS-ElevatedAdmins"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton03.Checked) {
                $Role = "ROLE-SCTS-RemoteDesktop"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton04.Checked) {
                $Role = "ROLE-SCTS-ServerAdmins"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton05.Checked) {
                $Role = "ROLE-SCTS-Helpdesk"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton06.Checked) {
                $Role = "ROLE-SCTS-Field Services"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton07.Checked) {
                $Role = "ROLE-SCTS-UnifiedComms"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton08.Checked) {
                $Role = "ROLE-SCTS-SQLDBA"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton09.Checked) {
                $Role = " ROLE-SCTS-HelpdeskManagers"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton10.Checked) {
                $Role = "ROLE-SCTS-Architect"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton11.Checked) {
                $Role = "ROLE-SCTS-Audit"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton12.Checked) {
                $Role = "ROLE-SCTS-Developer"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton13.Checked) {
                $Role = "ROLE-SCTS-DeveloperManager"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton14.Checked) {
                $Role = "ROLE-SCTS-Tester"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton15.Checked) {
                $Role = "ROLE-SCTS-TesterManager"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton16.Checked) {
                $Role = "ROLE-SCTS-CyberOperations"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton17.Checked) {
                $Role = "ROLE-SCTS-CyberOperationsManager"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton18.Checked) {
                $Role = "ROLE-SCTS-WebDeveloper"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton19.Checked) {
                $Role = "ROLE-SCTS-Networks"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            elseif ($BoxRadioButton20.Checked) {
                $Role = "ROLE-SCTS-NetworksManager"
                $Description = $Role + " - " + $ticket
                AdminMake
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You need to select a Team", 'New Admin', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                Return MainForm
            }
        }
    }
}
function AdminMake {
    if (Get-ADUser -Filter { SamAccountName -eq $UserSamAccountNameadmin }) {
        [System.Windows.Forms.MessageBox]::Show("Admin account already exists", 'New Admin', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        Return MainForm
    }
    else {
        #New-AdUser -Name "$SelectedAdminUser" -SamAccountName "$UserSamAccountNameadmin" -GivenName "$UserFirstName" -Surname "$UserSurName" -DisplayName "$SelectedAdminUser" -UserPrincipalName "$UserSamAccountNameadmin@scotcourts.local" -Office $Office -Description $Description  -Path $OU -Enabled $True -ChangePasswordAtLogon $false  -AccountPassword (ConvertTo-SecureString "TTyy7788!1" -AsPlainText -force) -passThru
        #Add-ADGroupMember -Identity $Role -Members $UserSamAccountNameadmin
        Write-Output $Description
        Write-Output $Role
        Write-Output $UserFirstName
        Write-Output $UserSurName
        Write-Output $UserSamAccountNameadmin
        Write-Output $SelectedAdminUser
        Return MainForm
    }
}

function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $SelectManForm = New-Object System.Windows.Forms.Form
    $SelectManForm.width = 550
    $SelectManForm.height = 375
    $SelectManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $SelectManForm.MinimizeBox = $False
    $SelectManForm.MaximizeBox = $False
    $SelectManForm.FormBorderStyle = 'Fixed3D'
    $SelectManForm.Text = 'New Admin Creator V1.0'
    $SelectManForm.Icon = $Icon
    $SelectManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    $Logo = [System.Drawing.Image]::Fromfile('\\saufs01\IT\Enterprise Team\Usermanagement\icons\SCTS.png')
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width = $Logo.Size.Width
    $pictureBox.Height = $Logo.Size.Height
    $pictureBox.Image = $Logo
    $SelectManForm.controls.add($pictureBox)
    $SelectManFormtext1 = New-Object System.Windows.Forms.Label
    $SelectManFormtext1.Location = '20,120'
    $SelectManFormtext1.size = '500,50'
    $SelectManFormtext1.Text = "This script creates an admin account from selected user. `n Ticket number is NEEDED"
    $EmailBox1 = New-Object System.Windows.Forms.GroupBox
    $EmailBox1.Location = '10,175'
    $EmailBox1.size = '500,75'
    $EmailBox1.text = '1. Select a UserName from the dropdown lists:'
    $SelectManFormtextLabel1 = New-Object System.Windows.Forms.Label
    $SelectManFormtextLabel1.Location = '20,40'
    $SelectManFormtextLabel1.size = '100,20'
    $SelectManFormtextLabel1.Text = 'UserName:' 
    $SelectManFormNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $SelectManFormNameComboBox1.Location = '125,35'
    $SelectManFormNameComboBox1.Size = '350, 310'
    $SelectManFormNameComboBox1.AutoCompleteMode = 'Suggest'
    $SelectManFormNameComboBox1.AutoCompleteSource = 'ListItems'
    $SelectManFormNameComboBox1.Sorted = $true;
    $SelectManFormNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $SelectManFormNameComboBox1.DataSource = $UsernameList
    $SelectManFormNameComboBox1.add_SelectedIndexChanged( { $ChangeThisUser.Text = "$($SelectManFormNameComboBox1.SelectedItem.ToString())" })
    $SelectManFormtext2 = New-Object System.Windows.Forms.Label
    $SelectManFormtext2.Location = '20,275'
    $SelectManFormtext2.size = '75,150'
    $SelectManFormtext2.Text = 'Create:'
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
            $SelectManForm.Close()
            $SelectManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    # Add all the Form controls on one line 
    $SelectManForm.Controls.AddRange(@($EmailBox1, $ChangeThisUser, $SelectManFormtext1, $SelectManFormtext2, $OKButton, $CancelButton))
    # Add all the GroupBox controls on one line
    $EmailBox1.Controls.AddRange(@($SelectManFormtextLabel1, $SelectManFormNameComboBox1))
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $SelectManForm.AcceptButton = $OKButton
    $SelectManForm.CancelButton = $CancelButton
    # Activate the form
    $SelectManForm.Add_Shown( { $SelectManForm.Activate() })    
    # Get the results from the button click
    $Result = $SelectManForm.ShowDialog()
    # If the OK button is selected
    if ($Result -eq 'OK') {
        if ($ChangeThisUser.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Renamer.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            Return MainForm
        }
        $SelectedUser = $ChangeThisUser.text
        $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'SamAccountName'
        $UserFirstName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'GivenName'
        $UserSurName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'Surname'
        $UserSamAccountNameadmin = $UserSamAccountName + "admin"
        $SelectedAdminUser = "$SelectedUser - Admin"
        AdminType
    }
}
MainForm
