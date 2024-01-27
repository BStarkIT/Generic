# New User
# Author        Brian Stark
# Date          14/07/2022
# Proofed       
# Tested date   
# Version       5.00
# Purpose       to create new O365 user accounts
#
# Changes       
#               V5.00   13/07/2022  BS  Powershell 7 changes, new paths, random password & Copy details to clip board
#               V4.00   13/06/2022  BS  O365 completion    
#               V3.00   11/05/2022  BS  O365 update
#               V2.00   2022        BS  fix for CSV's, start conditions & updates
#               V1.11   09/06/2021  BS  Added catch for space in names
#               V1.1    09/06/2021  BS  Contractor options added
#               V1.01   15/04/2021  BS  OU change
#               V1.00   25/08/2020  BS  Inital version of script
#
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\New\$Date.txt" -append
$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\User.ico'
$version = '5.00'
$WinTitle = "New User v$version."
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$UserName = $env:username
if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
}
else {
    Function SCTSForm {
        $WinTitle = "New User v$version - SCTS"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\Office.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroups.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\description.csv"
        $Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $SCTSForm = New-Object System.Windows.Forms.Form
        $SCTSForm.width = 800
        $SCTSForm.height = 500
        $SCTSForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $SCTSForm.Controlbox = $false
        $SCTSForm.Icon = $Icon
        $SCTSForm.FormBorderStyle = 'Fixed3D'
        $SCTSForm.Text = $WinTitle
        $SCTSForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $SCTSFormBox1 = New-Object System.Windows.Forms.GroupBox
        $SCTSFormBox1.Location = '20,20'
        $SCTSFormBox1.size = '730,110'
        $SCTSFormBox1.text = '1. The User Will be created with the following settings:'
        $SCTSFormB1Text1 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text1.Location = '40,40'
        $SCTSFormB1Text1.size = '65,20'
        $SCTSFormB1Text1.Text = 'Name:' 
        $SCTSFormB1Text2 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text2.Location = '140,40'
        $SCTSFormB1Text2.size = '100,20'
        $SCTSFormB1Text2.ForeColor = 'Blue'
        $SCTSFormB1Text2.Text = $DisplayName
        $SCTSFormB1Text4 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text4.Location = '40,80'
        $SCTSFormB1Text4.size = '100,20'
        $SCTSFormB1Text4.Text = 'Email address:'
        $SCTSFormB1Text5 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text5.Location = '140,80'
        $SCTSFormB1Text5.size = '250,20'
        $SCTSFormB1Text5.ForeColor = 'Blue'
        $SCTSFormB1Text5.Text = $mail
        $SCTSFormB1Text6 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text6.Location = '500,40'
        $SCTSFormB1Text6.size = '75,20'
        $SCTSFormB1Text6.Text = 'Logon:'
        $SCTSFormB1Text7 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text7.Location = '575,40'
        $SCTSFormB1Text7.size = '100,20'
        $SCTSFormB1Text7.ForeColor = 'Blue'
        $SCTSFormB1Text7.Text = $tentativeSAM
        $SCTSFormB1Text8 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text8.Location = '500,70'
        $SCTSFormB1Text8.size = '75,20'
        $SCTSFormB1Text8.Text = 'Password:'
        $SCTSFormB1Text9 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text9.Location = '575,70'
        $SCTSFormB1Text9.size = '120,20'
        $SCTSFormB1Text9.ForeColor = 'Blue'
        $SCTSFormB1Text9.Text = $pass
        $SCTSFormBox2 = New-Object System.Windows.Forms.GroupBox
        $SCTSFormBox2.Location = '20,140'
        $SCTSFormBox2.size = '730,240'
        $SCTSFormBox2.text = "2. Select the new users AD details:"
        $SCTSFormB2Text1 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text1.Location = '20,35'
        $SCTSFormB2Text1.size = '350,40'
        $SCTSFormB2Text1.Text = "Office field in AD:" 
        $SCTSFormB2Text2 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text2.Location = '20,75'
        $SCTSFormB2Text2.size = '350,40'
        $SCTSFormB2Text2.Text = "Description field in AD:" 
        $SCTSFormB2Text3 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text3.Location = '20,115'
        $SCTSFormB2Text3.size = '350,40'
        $SCTSFormB2Text3.Text = "Distribution List field in AD:" 
        $SCTSFormB2Text4 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text4.Location = '20,155'
        $SCTSFormB2Text4.size = '370,40'
        $SCTSFormB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $SCTSFormB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox1.Location = '425,30'
        $SCTSFormB2comboBox1.Size = '250,40'
        $SCTSFormB2comboBox1.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox1.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox1.Sorted = $false;
        $SCTSFormB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox1.SelectedItem = $SCTSFormB2comboBox1.Items[0]
        $SCTSFormB2comboBox1.DataSource = $OfficeList.Office 
        $SCTSFormB2comboBox1.add_SelectedIndexChanged( { $SCTSFormB2comboBoxOfficeSelect.Text = "$($SCTSFormB2comboBox1.SelectedItem.ToString())" })
        $SCTSFormB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox2.Location = '425,70'
        $SCTSFormB2comboBox2.Size = '250,40'
        $SCTSFormB2comboBox2.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox2.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox2.Sorted = $false;
        $SCTSFormB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox2.SelectedItem = $SCTSFormB2comboBox2.Items[0]
        $SCTSFormB2comboBox2.DataSource = $DescriptionList.Description 
        $SCTSFormB2comboBox2.add_SelectedIndexChanged( { $SCTSFormB2comboBoxDescriptionSelect.Text = "$($SCTSFormB2comboBox2.SelectedItem.ToString())" })
        $SCTSFormB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox3.Location = '425,110'
        $SCTSFormB2comboBox3.Size = '250,40'
        $SCTSFormB2comboBox3.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox3.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox3.Sorted = $true;
        $SCTSFormB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox3.SelectedItem = $SCTSFormB2comboBox3.Items[0]
        $SCTSFormB2comboBox3.DataSource = $Distributionlists
        $SCTSFormB2comboBox3.add_SelectedIndexChanged( { $SCTSFormB2comboBoxDistributionSelect.Text = "$($SCTSFormB2comboBox3.SelectedItem.ToString())" })
        $SCTSFormB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox4.Location = '425,150'
        $SCTSFormB2comboBox4.Size = '250,40'
        $SCTSFormB2comboBox4.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox4.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox4.Sorted = $false;
        $SCTSFormB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox4.SelectedItem = $SCTSFormB2comboBox4.Items[0]
        $SCTSFormB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $SCTSFormB2comboBox4.add_SelectedIndexChanged( { $SCTSFormB2comboBoxSecuritySelect.Text = "$($SCTSFormB2comboBox4.SelectedItem.ToString())" })
        $SCTSFormB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxOfficeSelect.Location = '20,600'
        $SCTSFormB2comboBoxOfficeSelect.size = '350,40'
        $SCTSFormB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxDescriptionSelect.Location = '20,650'
        $SCTSFormB2comboBoxDescriptionSelect.size = '350,40'
        $SCTSFormB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxDistributionSelect.Location = '20,700'
        $SCTSFormB2comboBoxDistributionSelect.size = '350,40'
        $SCTSFormB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxSecuritySelect.Location = '20,750'
        $SCTSFormB2comboBoxSecuritySelect.size = '350,40'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,400'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,400'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $SCTSForm.Close()
                $SCTSForm.Dispose()
                Return MainForm })
        $SCTSFormBox1.Controls.AddRange(@($SCTSFormB1Text1, $SCTSFormB1Text2, $SCTSFormB1Text4, $SCTSFormB1Text5, $SCTSFormB1Text6, $SCTSFormB1Text7, $SCTSFormB1Text8, $SCTSFormB1Text9))
        $SCTSFormBox2.Controls.AddRange(@($SCTSFormB2Text1, $SCTSFormB2Text2, $SCTSFormB2Text3, $SCTSFormB2Text4, $SCTSFormB2comboBox1, $SCTSFormB2comboBox2, $SCTSFormB2comboBox3, $SCTSFormB2comboBox4, $SCTSFormB2comboBoxOfficeSelect, $SCTSFormB2comboBoxDescriptionSelect , $SCTSFormB2comboBoxDistributionSelect, $SCTSFormB2comboBoxSecuritySelect))
        $SCTSForm.Controls.AddRange(@($SCTSFormBox1, $SCTSFormBox2, $OKButton, $CancelButton))
        $SCTSForm.AcceptButton = $OKButton
        $SCTSForm.CancelButton = $CancelButton
        $SCTSForm.Add_Shown( { $SCTSForm.Activate() })    
        $dialogResult = $SCTSForm.ShowDialog()
        if ($dialogResult -eq 'OK') {
            $Office = $SCTSFormB2comboBoxOfficeSelect.text
            $Description = $SCTSFormB2comboBoxDescriptionSelect.Text
            $DistributionGroup = $SCTSFormB2comboBoxDistributionSelect.Text
            $SecurityGroup = $SCTSFormB2comboBoxSecuritySelect.Text   
            $emailPrim = "SMTP:$mail"
            $emailProx = "$tentativeSAM@ScotCourtsTribunals.gov.uk"
            $SCTSForm.Close()
            $SCTSForm.Dispose()
            SBuildForm
        }
    }
    Function SCTSContForm {
        $WinTitle = "New User v$version - SCTS Contractor"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\Office.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroups.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\description.csv"
        $Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $SCTSForm = New-Object System.Windows.Forms.Form
        $SCTSForm.width = 800
        $SCTSForm.height = 650
        $SCTSForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $SCTSForm.Controlbox = $false
        $SCTSForm.Icon = $Icon
        $SCTSForm.FormBorderStyle = 'Fixed3D'
        $SCTSForm.Text = $WinTitle
        $SCTSForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $SCTSFormBox1 = New-Object System.Windows.Forms.GroupBox
        $SCTSFormBox1.Location = '20,20'
        $SCTSFormBox1.size = '730,110'
        $SCTSFormBox1.text = '1. The User Will be created with the following settings:'
        $SCTSFormB1Text1 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text1.Location = '40,40'
        $SCTSFormB1Text1.size = '65,20'
        $SCTSFormB1Text1.Text = 'Name:' 
        $SCTSFormB1Text2 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text2.Location = '140,40'
        $SCTSFormB1Text2.size = '100,20'
        $SCTSFormB1Text2.ForeColor = 'Blue'
        $SCTSFormB1Text2.Text = $DisplayName
        $SCTSFormB1Text4 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text4.Location = '40,80'
        $SCTSFormB1Text4.size = '100,20'
        $SCTSFormB1Text4.Text = 'Email address:'
        $SCTSFormB1Text5 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text5.Location = '140,80'
        $SCTSFormB1Text5.size = '250,20'
        $SCTSFormB1Text5.ForeColor = 'Blue'
        $SCTSFormB1Text5.Text = $mail
        $SCTSFormB1Text6 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text6.Location = '500,40'
        $SCTSFormB1Text6.size = '75,20'
        $SCTSFormB1Text6.Text = 'Logon:'
        $SCTSFormB1Text7 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text7.Location = '575,40'
        $SCTSFormB1Text7.size = '100,20'
        $SCTSFormB1Text7.ForeColor = 'Blue'
        $SCTSFormB1Text7.Text = $tentativeSAM
        $SCTSFormB1Text8 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text8.Location = '500,70'
        $SCTSFormB1Text8.size = '75,20'
        $SCTSFormB1Text8.Text = 'Password:'
        $SCTSFormB1Text9 = New-Object System.Windows.Forms.Label
        $SCTSFormB1Text9.Location = '575,70'
        $SCTSFormB1Text9.size = '120,20'
        $SCTSFormB1Text9.ForeColor = 'Blue'
        $SCTSFormB1Text9.Text = $pass
        $SCTSFormBox2 = New-Object System.Windows.Forms.GroupBox
        $SCTSFormBox2.Location = '20,140'
        $SCTSFormBox2.size = '730,200'
        $SCTSFormBox2.text = "2. Select the new users AD details:"
        $SCTSFormB2Text1 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text1.Location = '20,35'
        $SCTSFormB2Text1.size = '350,40'
        $SCTSFormB2Text1.Text = "Office field in AD:" 
        $SCTSFormB2Text2 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text2.Location = '20,75'
        $SCTSFormB2Text2.size = '350,40'
        $SCTSFormB2Text2.Text = "Description field in AD:" 
        $SCTSFormB2Text3 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text3.Location = '20,115'
        $SCTSFormB2Text3.size = '350,40'
        $SCTSFormB2Text3.Text = "Distribution List field in AD:" 
        $SCTSFormB2Text4 = New-Object System.Windows.Forms.Label
        $SCTSFormB2Text4.Location = '20,155'
        $SCTSFormB2Text4.size = '370,40'
        $SCTSFormB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $SCTSFormB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox1.Location = '425,30'
        $SCTSFormB2comboBox1.Size = '250,40'
        $SCTSFormB2comboBox1.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox1.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox1.Sorted = $false;
        $SCTSFormB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox1.SelectedItem = $SCTSFormB2comboBox1.Items[0]
        $SCTSFormB2comboBox1.DataSource = $OfficeList.Office 
        $SCTSFormB2comboBox1.add_SelectedIndexChanged( { $SCTSFormB2comboBoxOfficeSelect.Text = "$($SCTSFormB2comboBox1.SelectedItem.ToString())" })
        $SCTSFormB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox2.Location = '425,70'
        $SCTSFormB2comboBox2.Size = '250,40'
        $SCTSFormB2comboBox2.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox2.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox2.Sorted = $false;
        $SCTSFormB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox2.SelectedItem = $SCTSFormB2comboBox2.Items[0]
        $SCTSFormB2comboBox2.DataSource = $DescriptionList.Description 
        $SCTSFormB2comboBox2.add_SelectedIndexChanged( { $SCTSFormB2comboBoxDescriptionSelect.Text = "$($SCTSFormB2comboBox2.SelectedItem.ToString())" })
        $SCTSFormB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox3.Location = '425,110'
        $SCTSFormB2comboBox3.Size = '250,40'
        $SCTSFormB2comboBox3.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox3.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox3.Sorted = $true;
        $SCTSFormB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox3.SelectedItem = $SCTSFormB2comboBox3.Items[0]
        $SCTSFormB2comboBox3.DataSource = $Distributionlists
        $SCTSFormB2comboBox3.add_SelectedIndexChanged( { $SCTSFormB2comboBoxDistributionSelect.Text = "$($SCTSFormB2comboBox3.SelectedItem.ToString())" })
        $SCTSFormB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $SCTSFormB2comboBox4.Location = '425,150'
        $SCTSFormB2comboBox4.Size = '250,40'
        $SCTSFormB2comboBox4.AutoCompleteMode = 'Suggest'
        $SCTSFormB2comboBox4.AutoCompleteSource = 'ListItems'
        $SCTSFormB2comboBox4.Sorted = $false;
        $SCTSFormB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SCTSFormB2comboBox4.SelectedItem = $SCTSFormB2comboBox4.Items[0]
        $SCTSFormB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $SCTSFormB2comboBox4.add_SelectedIndexChanged( { $SCTSFormB2comboBoxSecuritySelect.Text = "$($SCTSFormB2comboBox4.SelectedItem.ToString())" })
        $SCTSFormB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxOfficeSelect.Location = '20,600'
        $SCTSFormB2comboBoxOfficeSelect.size = '350,40'
        $SCTSFormB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxDescriptionSelect.Location = '20,650'
        $SCTSFormB2comboBoxDescriptionSelect.size = '350,40'
        $SCTSFormB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxDistributionSelect.Location = '20,700'
        $SCTSFormB2comboBoxDistributionSelect.size = '350,40'
        $SCTSFormB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $SCTSFormB2comboBoxSecuritySelect.Location = '20,750'
        $SCTSFormB2comboBoxSecuritySelect.size = '350,40'
        $SCTSFormBox3 = New-Object System.Windows.Forms.GroupBox
        $SCTSFormBox3.Location = '20,350'
        $SCTSFormBox3.size = '260,240'
        $SCTSFormBox3.text = "3. Contractor end date:"
        $calendar = New-Object Windows.Forms.MonthCalendar -Property @{
            ShowTodayCircle   = $false
            MaxSelectionCount = 1
        }
        $calendar.Location = '25,25' 
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,450'
        $OKButton.Size = '100,60'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,450'
        $CancelButton.Size = '100,60'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $SCTSForm.Close()
                $SCTSForm.Dispose()
                Return MainForm })
        $SCTSFormBox1.Controls.AddRange(@($SCTSFormB1Text1, $SCTSFormB1Text2, $SCTSFormB1Text4, $SCTSFormB1Text5, $SCTSFormB1Text6, $SCTSFormB1Text7, $SCTSFormB1Text8, $SCTSFormB1Text9))
        $SCTSFormBox2.Controls.AddRange(@($SCTSFormB2Text1, $SCTSFormB2Text2, $SCTSFormB2Text3, $SCTSFormB2Text4, $SCTSFormB2comboBox1, $SCTSFormB2comboBox2, $SCTSFormB2comboBox3, $SCTSFormB2comboBox4, $SCTSFormB2comboBoxOfficeSelect, $SCTSFormB2comboBoxDescriptionSelect , $SCTSFormB2comboBoxDistributionSelect, $SCTSFormB2comboBoxSecuritySelect))
        $SCTSFormBox3.Controls.AddRange(@($calendar))
        $SCTSForm.Controls.AddRange(@($SCTSFormBox1, $SCTSFormBox2, $SCTSFormBox3, $OKButton, $CancelButton))
        $SCTSForm.AcceptButton = $OKButton
        $SCTSForm.CancelButton = $CancelButton
        $SCTSForm.Add_Shown( { $SCTSForm.Activate() })    
        $dialogResult = $SCTSForm.ShowDialog()
        $date = $calendar.SelectionStart
        $today = Get-Date
        if ($dialogResult -eq 'OK') {
            Write-Host "Selected $date"
            $Office = $SCTSFormB2comboBoxOfficeSelect.text
            $Description1 = $SCTSFormB2comboBoxDescriptionSelect.Text
            $Description = "Contractor - $Description1"
            $DistributionGroup = $SCTSFormB2comboBoxDistributionSelect.Text
            $SecurityGroup = $SCTSFormB2comboBoxSecuritySelect.Text   
            $emailPrim = "SMTP:$mail"
            $emailProx = "$tentativeSAM@ScotCourtsTribunals.gov.uk"
            $SCTSForm.Close()
            $SCTSForm.Dispose()
            SCBuildForm 
        
        }
    }
    Function Tribsform {
        $WinTitle = "New User v$version - Tribunals"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\OfficeTribs.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroupsTribs.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\DescriptionTribs.csv"
        $Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $Tribsform = New-Object System.Windows.Forms.Form
        $Tribsform.width = 800
        $Tribsform.height = 500
        $Tribsform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $Tribsform.Controlbox = $false
        $Tribsform.Icon = $Icon
        $Tribsform.FormBorderStyle = 'Fixed3D'
        $Tribsform.Text = $WinTitle
        $Tribsform.Font = New-Object System.Drawing.Font('Ariel', 10)
        $TribsformBox1 = New-Object System.Windows.Forms.GroupBox
        $TribsformBox1.Location = '20,20'
        $TribsformBox1.size = '730,110'
        $TribsformBox1.text = '1. The User Will be created with the following settings:'
        $TribsformB1Text1 = New-Object System.Windows.Forms.Label
        $TribsformB1Text1.Location = '40,40'
        $TribsformB1Text1.size = '65,20'
        $TribsformB1Text1.Text = 'Name:' 
        $TribsformB1Text2 = New-Object System.Windows.Forms.Label
        $TribsformB1Text2.Location = '140,40'
        $TribsformB1Text2.size = '100,20'
        $TribsformB1Text2.ForeColor = 'Blue'
        $TribsformB1Text2.Text = $DisplayName
        $TribsformB1Text4 = New-Object System.Windows.Forms.Label
        $TribsformB1Text4.Location = '40,80'
        $TribsformB1Text4.size = '100,20'
        $TribsformB1Text4.Text = 'Email address:'
        $TribsformB1Text5 = New-Object System.Windows.Forms.Label
        $TribsformB1Text5.Location = '140,80'
        $TribsformB1Text5.size = '280,20'
        $TribsformB1Text5.ForeColor = 'Blue'
        $TribsformB1Text5.Text = $mail
        $TribsformB1Text6 = New-Object System.Windows.Forms.Label
        $TribsformB1Text6.Location = '500,40'
        $TribsformB1Text6.size = '75,20'
        $TribsformB1Text6.Text = 'Logon:'
        $TribsformB1Text7 = New-Object System.Windows.Forms.Label
        $TribsformB1Text7.Location = '575,40'
        $TribsformB1Text7.size = '100,20'
        $TribsformB1Text7.ForeColor = 'Blue'
        $TribsformB1Text7.Text = $tentativeSAM
        $TribsformB1Text8 = New-Object System.Windows.Forms.Label
        $TribsformB1Text8.Location = '500,70'
        $TribsformB1Text8.size = '75,20'
        $TribsformB1Text8.Text = 'Password:'
        $TribsformB1Text9 = New-Object System.Windows.Forms.Label
        $TribsformB1Text9.Location = '575,70'
        $TribsformB1Text9.size = '120,20'
        $TribsformB1Text9.ForeColor = 'Blue'
        $TribsformB1Text9.Text = $pass
        $TribsformBox2 = New-Object System.Windows.Forms.GroupBox
        $TribsformBox2.Location = '20,140'
        $TribsformBox2.size = '730,240'
        $TribsformBox2.text = "2. Select the new users AD details:"
        $TribsformB2Text1 = New-Object System.Windows.Forms.Label
        $TribsformB2Text1.Location = '20,35'
        $TribsformB2Text1.size = '350,40'
        $TribsformB2Text1.Text = "Office field in AD:" 
        $TribsformB2Text2 = New-Object System.Windows.Forms.Label
        $TribsformB2Text2.Location = '20,75'
        $TribsformB2Text2.size = '350,40'
        $TribsformB2Text2.Text = "Description field in AD:" 
        $TribsformB2Text3 = New-Object System.Windows.Forms.Label
        $TribsformB2Text3.Location = '20,115'
        $TribsformB2Text3.size = '350,40'
        $TribsformB2Text3.Text = "Distribution List field in AD:" 
        $TribsformB2Text4 = New-Object System.Windows.Forms.Label
        $TribsformB2Text4.Location = '20,155'
        $TribsformB2Text4.size = '370,40'
        $TribsformB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $TribsformB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox1.Location = '425,30'
        $TribsformB2comboBox1.Size = '250,40'
        $TribsformB2comboBox1.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox1.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox1.Sorted = $false;
        $TribsformB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox1.SelectedItem = $TribsformB2comboBox1.Items[0]
        $TribsformB2comboBox1.DataSource = $OfficeList.Office 
        $TribsformB2comboBox1.add_SelectedIndexChanged( { $TribsformB2comboBoxOfficeSelect.Text = "$($TribsformB2comboBox1.SelectedItem.ToString())" })
        $TribsformB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox2.Location = '425,70'
        $TribsformB2comboBox2.Size = '250,40'
        $TribsformB2comboBox2.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox2.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox2.Sorted = $false;
        $TribsformB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox2.SelectedItem = $TribsformB2comboBox2.Items[0]
        $TribsformB2comboBox2.DataSource = $DescriptionList.Description 
        $TribsformB2comboBox2.add_SelectedIndexChanged( { $TribsformB2comboBoxDescriptionSelect.Text = "$($TribsformB2comboBox2.SelectedItem.ToString())" })
        $TribsformB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox3.Location = '425,110'
        $TribsformB2comboBox3.Size = '250,40'
        $TribsformB2comboBox3.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox3.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox3.Sorted = $true;
        $TribsformB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox3.SelectedItem = $TribsformB2comboBox3.Items[0]
        $TribsformB2comboBox3.DataSource = $Distributionlists
        $TribsformB2comboBox3.add_SelectedIndexChanged( { $TribsformB2comboBoxDistributionSelect.Text = "$($TribsformB2comboBox3.SelectedItem.ToString())" })
        $TribsformB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox4.Location = '425,150'
        $TribsformB2comboBox4.Size = '250,40'
        $TribsformB2comboBox4.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox4.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox4.Sorted = $false;
        $TribsformB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox4.SelectedItem = $TribsformB2comboBox4.Items[0]
        $TribsformB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $TribsformB2comboBox4.add_SelectedIndexChanged( { $TribsformB2comboBoxSecuritySelect.Text = "$($TribsformB2comboBox4.SelectedItem.ToString())" })
        $TribsformB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxOfficeSelect.Location = '20,600'
        $TribsformB2comboBoxOfficeSelect.size = '350,40'
        $TribsformB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxDescriptionSelect.Location = '20,650'
        $TribsformB2comboBoxDescriptionSelect.size = '350,40'
        $TribsformB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxDistributionSelect.Location = '20,700'
        $TribsformB2comboBoxDistributionSelect.size = '350,40'
        $TribsformB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxSecuritySelect.Location = '20,750'
        $TribsformB2comboBoxSecuritySelect.size = '350,40'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,400'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,400'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $Tribsform.Close()
                $Tribsform.Dispose()
                Return MainForm })
        $TribsformBox1.Controls.AddRange(@($TribsformB1Text1, $TribsformB1Text2, $TribsformB1Text4, $TribsformB1Text5, $TribsformB1Text6, $TribsformB1Text7, $TribsformB1Text8, $TribsformB1Text9))
        $TribsformBox2.Controls.AddRange(@($TribsformB2Text1, $TribsformB2Text2, $TribsformB2Text3, $TribsformB2Text4, $TribsformB2comboBox1, $TribsformB2comboBox2, $TribsformB2comboBox3, $TribsformB2comboBox4, $TribsformB2comboBoxOfficeSelect, $TribsformB2comboBoxDescriptionSelect , $TribsformB2comboBoxDistributionSelect, $TribsformB2comboBoxSecuritySelect))
        $Tribsform.Controls.AddRange(@($TribsformBox1, $TribsformBox2, $OKButton, $CancelButton))
        $Tribsform.AcceptButton = $OKButton
        $Tribsform.CancelButton = $CancelButton
        $Tribsform.Add_Shown( { $Tribsform.Activate() })    
        $dialogResult = $Tribsform.ShowDialog()
        if ($dialogResult -eq 'OK') {
            $Office = $TribsformB2comboBoxOfficeSelect.text
            $Description = $TribsformB2comboBoxDescriptionSelect.Text
            $DistributionGroup = $TribsformB2comboBoxDistributionSelect.Text
            $SecurityGroup = $TribsformB2comboBoxSecuritySelect.Text   
            $emailPrim = "SMTP:$mail"
            $emailProx = "$tentativeSAM@scotcourts.gov.uk"
            $SAM = $tentativeSAM
            $Tribsform.Close()
            $Tribsform.Dispose()
            TBuildForm
        }
    }
    Function TribsContform {
        $WinTitle = "New User v$version - Contractor - Tribunals"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\OfficeTribs.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroupsTribs.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\DescriptionTribs.csv"
        $Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $Tribsform = New-Object System.Windows.Forms.Form
        $Tribsform.width = 800
        $Tribsform.height = 500
        $Tribsform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $Tribsform.Controlbox = $false
        $Tribsform.Icon = $Icon
        $Tribsform.FormBorderStyle = 'Fixed3D'
        $Tribsform.Text = $WinTitle
        $Tribsform.Font = New-Object System.Drawing.Font('Ariel', 10)
        $TribsformBox1 = New-Object System.Windows.Forms.GroupBox
        $TribsformBox1.Location = '20,20'
        $TribsformBox1.size = '730,110'
        $TribsformBox1.text = '1. The User Will be created with the following settings:'
        $TribsformB1Text1 = New-Object System.Windows.Forms.Label
        $TribsformB1Text1.Location = '40,40'
        $TribsformB1Text1.size = '65,20'
        $TribsformB1Text1.Text = 'Name:' 
        $TribsformB1Text2 = New-Object System.Windows.Forms.Label
        $TribsformB1Text2.Location = '140,40'
        $TribsformB1Text2.size = '100,20'
        $TribsformB1Text2.ForeColor = 'Blue'
        $TribsformB1Text2.Text = $DisplayName
        $TribsformB1Text4 = New-Object System.Windows.Forms.Label
        $TribsformB1Text4.Location = '40,80'
        $TribsformB1Text4.size = '100,20'
        $TribsformB1Text4.Text = 'Email address:'
        $TribsformB1Text5 = New-Object System.Windows.Forms.Label
        $TribsformB1Text5.Location = '140,80'
        $TribsformB1Text5.size = '280,20'
        $TribsformB1Text5.ForeColor = 'Blue'
        $TribsformB1Text5.Text = $mail
        $TribsformB1Text6 = New-Object System.Windows.Forms.Label
        $TribsformB1Text6.Location = '500,40'
        $TribsformB1Text6.size = '75,20'
        $TribsformB1Text6.Text = 'Logon:'
        $TribsformB1Text7 = New-Object System.Windows.Forms.Label
        $TribsformB1Text7.Location = '575,40'
        $TribsformB1Text7.size = '100,20'
        $TribsformB1Text7.ForeColor = 'Blue'
        $TribsformB1Text7.Text = $tentativeSAM
        $TribsformB1Text8 = New-Object System.Windows.Forms.Label
        $TribsformB1Text8.Location = '500,70'
        $TribsformB1Text8.size = '75,20'
        $TribsformB1Text8.Text = 'Password:'
        $TribsformB1Text9 = New-Object System.Windows.Forms.Label
        $TribsformB1Text9.Location = '575,70'
        $TribsformB1Text9.size = '120,20'
        $TribsformB1Text9.ForeColor = 'Blue'
        $TribsformB1Text9.Text = $pass
        $TribsformBox2 = New-Object System.Windows.Forms.GroupBox
        $TribsformBox2.Location = '20,140'
        $TribsformBox2.size = '730,240'
        $TribsformBox2.text = "2. Select the new users AD details:"
        $TribsformB2Text1 = New-Object System.Windows.Forms.Label
        $TribsformB2Text1.Location = '20,35'
        $TribsformB2Text1.size = '350,40'
        $TribsformB2Text1.Text = "Office field in AD:" 
        $TribsformB2Text2 = New-Object System.Windows.Forms.Label
        $TribsformB2Text2.Location = '20,75'
        $TribsformB2Text2.size = '350,40'
        $TribsformB2Text2.Text = "Description field in AD:" 
        $TribsformB2Text3 = New-Object System.Windows.Forms.Label
        $TribsformB2Text3.Location = '20,115'
        $TribsformB2Text3.size = '350,40'
        $TribsformB2Text3.Text = "Distribution List field in AD:" 
        $TribsformB2Text4 = New-Object System.Windows.Forms.Label
        $TribsformB2Text4.Location = '20,155'
        $TribsformB2Text4.size = '370,40'
        $TribsformB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $TribsformB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox1.Location = '425,30'
        $TribsformB2comboBox1.Size = '250,40'
        $TribsformB2comboBox1.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox1.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox1.Sorted = $false;
        $TribsformB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox1.SelectedItem = $TribsformB2comboBox1.Items[0]
        $TribsformB2comboBox1.DataSource = $OfficeList.Office 
        $TribsformB2comboBox1.add_SelectedIndexChanged( { $TribsformB2comboBoxOfficeSelect.Text = "$($TribsformB2comboBox1.SelectedItem.ToString())" })
        $TribsformB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox2.Location = '425,70'
        $TribsformB2comboBox2.Size = '250,40'
        $TribsformB2comboBox2.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox2.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox2.Sorted = $false;
        $TribsformB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox2.SelectedItem = $TribsformB2comboBox2.Items[0]
        $TribsformB2comboBox2.DataSource = $DescriptionList.Description 
        $TribsformB2comboBox2.add_SelectedIndexChanged( { $TribsformB2comboBoxDescriptionSelect.Text = "$($TribsformB2comboBox2.SelectedItem.ToString())" })
        $TribsformB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox3.Location = '425,110'
        $TribsformB2comboBox3.Size = '250,40'
        $TribsformB2comboBox3.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox3.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox3.Sorted = $true;
        $TribsformB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox3.SelectedItem = $TribsformB2comboBox3.Items[0]
        $TribsformB2comboBox3.DataSource = $Distributionlists
        $TribsformB2comboBox3.add_SelectedIndexChanged( { $TribsformB2comboBoxDistributionSelect.Text = "$($TribsformB2comboBox3.SelectedItem.ToString())" })
        $TribsformB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $TribsformB2comboBox4.Location = '425,150'
        $TribsformB2comboBox4.Size = '250,40'
        $TribsformB2comboBox4.AutoCompleteMode = 'Suggest'
        $TribsformB2comboBox4.AutoCompleteSource = 'ListItems'
        $TribsformB2comboBox4.Sorted = $false;
        $TribsformB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $TribsformB2comboBox4.SelectedItem = $TribsformB2comboBox4.Items[0]
        $TribsformB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $TribsformB2comboBox4.add_SelectedIndexChanged( { $TribsformB2comboBoxSecuritySelect.Text = "$($TribsformB2comboBox4.SelectedItem.ToString())" })
        $TribsformB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxOfficeSelect.Location = '20,600'
        $TribsformB2comboBoxOfficeSelect.size = '350,40'
        $TribsformB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxDescriptionSelect.Location = '20,650'
        $TribsformB2comboBoxDescriptionSelect.size = '350,40'
        $TribsformB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxDistributionSelect.Location = '20,700'
        $TribsformB2comboBoxDistributionSelect.size = '350,40'
        $TribsformB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $TribsformB2comboBoxSecuritySelect.Location = '20,750'
        $TribsformB2comboBoxSecuritySelect.size = '350,40'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,400'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,400'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $Tribsform.Close()
                $Tribsform.Dispose()
                Return MainForm })
        $TribsformBox1.Controls.AddRange(@($TribsformB1Text1, $TribsformB1Text2, $TribsformB1Text4, $TribsformB1Text5, $TribsformB1Text6, $TribsformB1Text7, $TribsformB1Text8, $TribsformB1Text9))
        $TribsformBox2.Controls.AddRange(@($TribsformB2Text1, $TribsformB2Text2, $TribsformB2Text3, $TribsformB2Text4, $TribsformB2comboBox1, $TribsformB2comboBox2, $TribsformB2comboBox3, $TribsformB2comboBox4, $TribsformB2comboBoxOfficeSelect, $TribsformB2comboBoxDescriptionSelect , $TribsformB2comboBoxDistributionSelect, $TribsformB2comboBoxSecuritySelect))
        $Tribsform.Controls.AddRange(@($TribsformBox1, $TribsformBox2, $OKButton, $CancelButton))
        $Tribsform.AcceptButton = $OKButton
        $Tribsform.CancelButton = $CancelButton
        $Tribsform.Add_Shown( { $Tribsform.Activate() })    
        $dialogResult = $Tribsform.ShowDialog()
        if ($dialogResult -eq 'OK') {
            if ($date = $today ) {
                [System.Windows.Forms.MessageBox]::Show("You need to set an end date!", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            else {
                $Office = $TribsformB2comboBoxOfficeSelect.text
                $Description = $TribsformB2comboBoxDescriptionSelect.Text
                $DistributionGroup = $TribsformB2comboBoxDistributionSelect.Text
                $SecurityGroup = $TribsformB2comboBoxSecuritySelect.Text   
                $emailPrim = "SMTP:$mail"
                $emailProx = "$tentativeSAM@scotcourts.gov.uk"
                $SAM = $tentativeSAM
                $Tribsform.Close()
                $Tribsform.Dispose()
                TCBuildForm
            }
        }
    }
    Function Judform {
        $WinTitle = "New User v$version - Judiciary"
        $mail = "$tentativeSAM@ScotCourts.gov.uk"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\JudicialUsers\officeJud.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroups.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\JudicialUsers\descriptionJud.csv"
        $Distributionlists = Import-Csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\distributionJud.csv"
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $Judform = New-Object System.Windows.Forms.Form
        $Judform.width = 800
        $Judform.height = 500
        $Judform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $Judform.Controlbox = $false
        $Judform.Icon = $Icon
        $Judform.FormBorderStyle = 'Fixed3D'
        $Judform.Text = $WinTitle
        $Judform.Font = New-Object System.Drawing.Font('Ariel', 10)
        $JudformBox1 = New-Object System.Windows.Forms.GroupBox
        $JudformBox1.Location = '20,20'
        $JudformBox1.size = '730,110'
        $JudformBox1.text = '1. The User Will be created with the following settings:'
        $JudformB1Text1 = New-Object System.Windows.Forms.Label
        $JudformB1Text1.Location = '40,40'
        $JudformB1Text1.size = '65,20'
        $JudformB1Text1.Text = 'Name:' 
        $JudformB1Text2 = New-Object System.Windows.Forms.Label
        $JudformB1Text2.Location = '140,40'
        $JudformB1Text2.size = '200,20'
        $JudformB1Text2.ForeColor = 'Blue'
        $JudformB1Text2.Text = $DisplayName
        $JudformB1Text4 = New-Object System.Windows.Forms.Label
        $JudformB1Text4.Location = '40,80'
        $JudformB1Text4.size = '100,20'
        $JudformB1Text4.Text = 'Email address:'
        $JudformB1Text5 = New-Object System.Windows.Forms.Label
        $JudformB1Text5.Location = '140,80'
        $JudformB1Text5.size = '250,20'
        $JudformB1Text5.ForeColor = 'Blue'
        $JudformB1Text5.Text = $mail
        $JudformB1Text6 = New-Object System.Windows.Forms.Label
        $JudformB1Text6.Location = '500,40'
        $JudformB1Text6.size = '75,20'
        $JudformB1Text6.Text = 'Logon:'
        $JudformB1Text7 = New-Object System.Windows.Forms.Label
        $JudformB1Text7.Location = '575,40'
        $JudformB1Text7.size = '100,20'
        $JudformB1Text7.ForeColor = 'Blue'
        $JudformB1Text7.Text = $tentativeSAM
        $JudformB1Text8 = New-Object System.Windows.Forms.Label
        $JudformB1Text8.Location = '500,70'
        $JudformB1Text8.size = '75,20'
        $JudformB1Text8.Text = 'Password:'
        $JudformB1Text9 = New-Object System.Windows.Forms.Label
        $JudformB1Text9.Location = '575,70'
        $JudformB1Text9.size = '120,20'
        $JudformB1Text9.ForeColor = 'Blue'
        $JudformB1Text9.Text = $pass
        $JudformBox2 = New-Object System.Windows.Forms.GroupBox
        $JudformBox2.Location = '20,140'
        $JudformBox2.size = '730,240'
        $JudformBox2.text = "2. Select the new users AD details:"
        $JudformB2Text1 = New-Object System.Windows.Forms.Label
        $JudformB2Text1.Location = '20,70'
        $JudformB2Text1.size = '350,30'
        $JudformB2Text1.Text = "Office field in AD:" 
        $JudformB2Text2 = New-Object System.Windows.Forms.Label
        $JudformB2Text2.Location = '20,105'
        $JudformB2Text2.size = '350,30'
        $JudformB2Text2.Text = "Description field in AD:" 
        $JudformB2Text3 = New-Object System.Windows.Forms.Label
        $JudformB2Text3.Location = '20,140'
        $JudformB2Text3.size = '350,30'
        $JudformB2Text3.Text = "Distribution List field in AD:" 
        $JudformB2Text4 = New-Object System.Windows.Forms.Label
        $JudformB2Text4.Location = '20,175'
        $JudformB2Text4.size = '370,30'
        $JudformB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $JudformB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox1.Location = '425,70'
        $JudformB2comboBox1.Size = '250,40'
        $JudformB2comboBox1.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox1.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox1.Sorted = $false;
        $JudformB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox1.SelectedItem = $JudformB2comboBox1.Items[0]
        $JudformB2comboBox1.DataSource = $OfficeList.Office 
        $JudformB2comboBox1.add_SelectedIndexChanged( { $JudformB2comboBoxOfficeSelect.Text = "$($JudformB2comboBox1.SelectedItem.ToString())" })
        $JudformB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox2.Location = '425,105'
        $JudformB2comboBox2.Size = '250,40'
        $JudformB2comboBox2.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox2.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox2.Sorted = $false;
        $JudformB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox2.SelectedItem = $JudformB2comboBox2.Items[0]
        $JudformB2comboBox2.DataSource = $DescriptionList.Description 
        $JudformB2comboBox2.add_SelectedIndexChanged( { $JudformB2comboBoxDescriptionSelect.Text = "$($JudformB2comboBox2.SelectedItem.ToString())" })
        $JudformB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox3.Location = '425,140'
        $JudformB2comboBox3.Size = '250,40'
        $JudformB2comboBox3.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox3.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox3.Sorted = $false;
        $JudformB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox3.SelectedItem = $JudformB2comboBox3.Items[0]
        $JudformB2comboBox3.DataSource = $Distributionlists.distribution
        $JudformB2comboBox3.add_SelectedIndexChanged( { $JudformB2comboBoxDistributionSelect.Text = "$($JudformB2comboBox3.SelectedItem.ToString())" })
        $JudformB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox4.Location = '425,175'
        $JudformB2comboBox4.Size = '250,40'
        $JudformB2comboBox4.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox4.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox4.Sorted = $false;
        $JudformB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox4.SelectedItem = $JudformB2comboBox4.Items[0]
        $JudformB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $JudformB2comboBox4.add_SelectedIndexChanged( { $JudformB2comboBoxSecuritySelect.Text = "$($JudformB2comboBox4.SelectedItem.ToString())" })
        $JudformB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxOfficeSelect.Location = '20,600'
        $JudformB2comboBoxOfficeSelect.size = '350,40'
        $JudformB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxDescriptionSelect.Location = '20,650'
        $JudformB2comboBoxDescriptionSelect.size = '350,40'
        $JudformB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxDistributionSelect.Location = '20,700'
        $JudformB2comboBoxDistributionSelect.size = '350,40'
        $JudformB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxSecuritySelect.Location = '20,750'
        $JudformB2comboBoxSecuritySelect.size = '350,40'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,400'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,400'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $Judform.Close()
                $Judform.Dispose()
                Return MainForm })
        $JudformBox1.Controls.AddRange(@($JudformB1Text1, $JudformB1Text2, $JudformB1Text4, $JudformB1Text5, $JudformB1Text6, $JudformB1Text7, $JudformB1Text8, $JudformB1Text9))
        $JudformBox2.Controls.AddRange(@($JudformB2Text1, $JudformB2Text2, $JudformB2Text3, $JudformB2Text4, $JudformB2comboBox1, $JudformB2comboBox2, $JudformB2comboBox3, $JudformB2comboBox4, $JudformB2comboBox5, $JudformB2comboBoxOfficeSelect, $JudformB2comboBoxDescriptionSelect , $JudformB2comboBoxDistributionSelect, $JudformB2comboBoxSecuritySelect))
        $Judform.Controls.AddRange(@($JudformBox1, $JudformBox2, $OKButton, $CancelButton))
        $Judform.AcceptButton = $OKButton
        $Judform.CancelButton = $CancelButton
        $Judform.Add_Shown( { $Judform.Activate() })    
        $dialogResult = $Judform.ShowDialog()
        if ($dialogResult -eq 'OK') {
            $Office = $JudformB2comboBoxOfficeSelect.text
            $Description = $JudformB2comboBoxDescriptionSelect.Text
            $DistributionGroup = $JudformB2comboBoxDistributionSelect.Text
            $SecurityGroup = $JudformB2comboBoxSecuritySelect.Text   
            $emailPrim = "SMTP:$mail"
            $emailProx = "$tentativeSAM@scotcourtstribunals.gov.uk"
            $SAM = $tentativeSAM
            $Judform.Close()
            $Judform.Dispose()
            JBuildForm
        }
    }
    Function JudContform {
        $WinTitle = "New User v$version - Contractor - Judiciary"
        $mail = "$tentativeSAM@ScotCourts.gov.uk"
        $OfficeList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\JudicialUsers\officeJud.csv"
        $SecurityGroupsList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\SecurityGroups.csv"
        $DescriptionList = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\JudicialUsers\descriptionJud.csv"
        $Distributionlists = Import-Csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\distributionJud.csv"
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $Judform = New-Object System.Windows.Forms.Form
        $Judform.width = 800
        $Judform.height = 500
        $Judform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $Judform.Controlbox = $false
        $Judform.Icon = $Icon
        $Judform.FormBorderStyle = 'Fixed3D'
        $Judform.Text = $WinTitle
        $Judform.Font = New-Object System.Drawing.Font('Ariel', 10)
        $JudformBox1 = New-Object System.Windows.Forms.GroupBox
        $JudformBox1.Location = '20,20'
        $JudformBox1.size = '730,110'
        $JudformBox1.text = '1. The User Will be created with the following settings:'
        $JudformB1Text1 = New-Object System.Windows.Forms.Label
        $JudformB1Text1.Location = '40,40'
        $JudformB1Text1.size = '65,20'
        $JudformB1Text1.Text = 'Name:' 
        $JudformB1Text2 = New-Object System.Windows.Forms.Label
        $JudformB1Text2.Location = '140,40'
        $JudformB1Text2.size = '200,20'
        $JudformB1Text2.ForeColor = 'Blue'
        $JudformB1Text2.Text = $DisplayName
        $JudformB1Text4 = New-Object System.Windows.Forms.Label
        $JudformB1Text4.Location = '40,80'
        $JudformB1Text4.size = '100,20'
        $JudformB1Text4.Text = 'Email address:'
        $JudformB1Text5 = New-Object System.Windows.Forms.Label
        $JudformB1Text5.Location = '140,80'
        $JudformB1Text5.size = '250,20'
        $JudformB1Text5.ForeColor = 'Blue'
        $JudformB1Text5.Text = $mail
        $JudformB1Text6 = New-Object System.Windows.Forms.Label
        $JudformB1Text6.Location = '500,40'
        $JudformB1Text6.size = '75,20'
        $JudformB1Text6.Text = 'Logon:'
        $JudformB1Text7 = New-Object System.Windows.Forms.Label
        $JudformB1Text7.Location = '575,40'
        $JudformB1Text7.size = '100,20'
        $JudformB1Text7.ForeColor = 'Blue'
        $JudformB1Text7.Text = $tentativeSAM
        $JudformB1Text8 = New-Object System.Windows.Forms.Label
        $JudformB1Text8.Location = '500,70'
        $JudformB1Text8.size = '75,20'
        $JudformB1Text8.Text = 'Password:'
        $JudformB1Text9 = New-Object System.Windows.Forms.Label
        $JudformB1Text9.Location = '575,70'
        $JudformB1Text9.size = '120,20'
        $JudformB1Text9.ForeColor = 'Blue'
        $JudformB1Text9.Text = $pass
        $JudformBox2 = New-Object System.Windows.Forms.GroupBox
        $JudformBox2.Location = '20,140'
        $JudformBox2.size = '730,240'
        $JudformBox2.text = "2. Select the new users AD details:"
        $JudformB2Text1 = New-Object System.Windows.Forms.Label
        $JudformB2Text1.Location = '20,70'
        $JudformB2Text1.size = '350,30'
        $JudformB2Text1.Text = "Office field in AD:" 
        $JudformB2Text2 = New-Object System.Windows.Forms.Label
        $JudformB2Text2.Location = '20,105'
        $JudformB2Text2.size = '350,30'
        $JudformB2Text2.Text = "Description field in AD:" 
        $JudformB2Text3 = New-Object System.Windows.Forms.Label
        $JudformB2Text3.Location = '20,140'
        $JudformB2Text3.size = '350,30'
        $JudformB2Text3.Text = "Distribution List field in AD:" 
        $JudformB2Text4 = New-Object System.Windows.Forms.Label
        $JudformB2Text4.Location = '20,175'
        $JudformB2Text4.size = '370,30'
        $JudformB2Text4.Text = "Security Group field in AD (to access p and s drives):" 
        $JudformB2comboBox1 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox1.Location = '425,70'
        $JudformB2comboBox1.Size = '250,40'
        $JudformB2comboBox1.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox1.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox1.Sorted = $false;
        $JudformB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox1.SelectedItem = $JudformB2comboBox1.Items[0]
        $JudformB2comboBox1.DataSource = $OfficeList.Office 
        $JudformB2comboBox1.add_SelectedIndexChanged( { $JudformB2comboBoxOfficeSelect.Text = "$($JudformB2comboBox1.SelectedItem.ToString())" })
        $JudformB2comboBox2 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox2.Location = '425,105'
        $JudformB2comboBox2.Size = '250,40'
        $JudformB2comboBox2.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox2.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox2.Sorted = $false;
        $JudformB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox2.SelectedItem = $JudformB2comboBox2.Items[0]
        $JudformB2comboBox2.DataSource = $DescriptionList.Description 
        $JudformB2comboBox2.add_SelectedIndexChanged( { $JudformB2comboBoxDescriptionSelect.Text = "$($JudformB2comboBox2.SelectedItem.ToString())" })
        $JudformB2comboBox3 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox3.Location = '425,140'
        $JudformB2comboBox3.Size = '250,40'
        $JudformB2comboBox3.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox3.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox3.Sorted = $false;
        $JudformB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox3.SelectedItem = $JudformB2comboBox3.Items[0]
        $JudformB2comboBox3.DataSource = $Distributionlists.distribution
        $JudformB2comboBox3.add_SelectedIndexChanged( { $JudformB2comboBoxDistributionSelect.Text = "$($JudformB2comboBox3.SelectedItem.ToString())" })
        $JudformB2comboBox4 = New-Object System.Windows.Forms.ComboBox
        $JudformB2comboBox4.Location = '425,175'
        $JudformB2comboBox4.Size = '250,40'
        $JudformB2comboBox4.AutoCompleteMode = 'Suggest'
        $JudformB2comboBox4.AutoCompleteSource = 'ListItems'
        $JudformB2comboBox4.Sorted = $false;
        $JudformB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudformB2comboBox4.SelectedItem = $JudformB2comboBox4.Items[0]
        $JudformB2comboBox4.DataSource = $SecurityGroupsList.Securitygroup 
        $JudformB2comboBox4.add_SelectedIndexChanged( { $JudformB2comboBoxSecuritySelect.Text = "$($JudformB2comboBox4.SelectedItem.ToString())" })
        $JudformB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxOfficeSelect.Location = '20,600'
        $JudformB2comboBoxOfficeSelect.size = '350,40'
        $JudformB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxDescriptionSelect.Location = '20,650'
        $JudformB2comboBoxDescriptionSelect.size = '350,40'
        $JudformB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxDistributionSelect.Location = '20,700'
        $JudformB2comboBoxDistributionSelect.size = '350,40'
        $JudformB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $JudformB2comboBoxSecuritySelect.Location = '20,750'
        $JudformB2comboBoxSecuritySelect.size = '350,40'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,400'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,400'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $Judform.Close()
                $Judform.Dispose()
                Return MainForm })
        $JudformBox1.Controls.AddRange(@($JudformB1Text1, $JudformB1Text2, $JudformB1Text4, $JudformB1Text5, $JudformB1Text6, $JudformB1Text7, $JudformB1Text8, $JudformB1Text9))
        $JudformBox2.Controls.AddRange(@($JudformB2Text1, $JudformB2Text2, $JudformB2Text3, $JudformB2Text4, $JudformB2comboBox1, $JudformB2comboBox2, $JudformB2comboBox3, $JudformB2comboBox4, $JudformB2comboBox5, $JudformB2comboBoxOfficeSelect, $JudformB2comboBoxDescriptionSelect , $JudformB2comboBoxDistributionSelect, $JudformB2comboBoxSecuritySelect))
        $Judform.Controls.AddRange(@($JudformBox1, $JudformBox2, $OKButton, $CancelButton))
        $Judform.AcceptButton = $OKButton
        $Judform.CancelButton = $CancelButton
        $Judform.Add_Shown( { $Judform.Activate() })    
        $dialogResult = $Judform.ShowDialog()
        if ($dialogResult -eq 'OK') {
            if ($date = $today ) {
                [System.Windows.Forms.MessageBox]::Show("You need to set an end date!", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            else {        
                $Office = $JudformB2comboBoxOfficeSelect.text
                $Description = $JudformB2comboBoxDescriptionSelect.Text
                $DistributionGroup = $JudformB2comboBoxDistributionSelect.Text
                $SecurityGroup = $JudformB2comboBoxSecuritySelect.Text   
                $emailPrim = "SMTP:$mail"
                $emailProx = "$tentativeSAM@scotcourtstribunals.gov.uk"
                $SAM = $tentativeSAM
                $Judform.Close()
                $Judform.Dispose()
                JCBuildForm
            }
        }
    }
    function SBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
        Set-ADUser -Identity $SAM  -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    function SCBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description  -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
        Set-ADUser -Identity $SAM  -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        Set-ADAccountExpiration -Identity $SAM -DateTime $date
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    function TBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" -Members $SAM
        Add-ADGroupMember -Identity "All Users Tribunals" -Members $SAM
        Add-ADGroupMember -Identity "acl_All STS Users_readwrite" -Members $SAM
        Add-ADGroupMember -Identity "acl_All_Tribunals_Users" -Members $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERT" }
        Set-ADUser -Identity $SAM  -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    function TCBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" -Members $SAM
        Add-ADGroupMember -Identity "All Users Tribunals" -Members $SAM
        Add-ADGroupMember -Identity "acl_All STS Users_readwrite" -Members $SAM
        Add-ADGroupMember -Identity "acl_All_Tribunals_Users" -Members $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERT" }
        Set-ADUser -Identity $SAM  -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        Set-ADAccountExpiration -Identity $SAM -DateTime $date
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    function JBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" -Members $SAM
        Add-ADGroupMember -Identity "All Judicial Studies Access" -Members $SAM
        Add-ADGroupMember -Identity "Judicial Studies Access" -Members $SAM
        Add-ADGroupMember -Identity "GPO SF - Judicial Hub Home Page" -Members $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
        Set-ADUser -Identity $SAM -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    function JCBuildForm {
        $OU = "OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server SAU-DC-04.scotcourts.local -AccountPassword $Password -passThru
        Add-Type -AssemblyName System.Windows.Forms 
        $objForm = New-Object System.Windows.Forms.Form
        $objForm.Text = $WinTitle
        $objForm.Size = New-Object System.Drawing.Size(450, 270)
        $objForm.StartPosition = "CenterScreen"
        $objForm.Controlbox = $false
        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
        $objLabel.Size = New-Object System.Drawing.Size(300, 120)
        $objLabel.Text = "A New User Account is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the SCTS Users OU.`n`nPlease Wait."
        $objForm.Controls.Add($objLabel)
        $objForm.Show() | Out-Null
        Start-Sleep -Seconds 5
        $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
        Enable-RemoteMailbox -Identity $SAM -RemoteRoutingAddress $RoutingAddress -DomainController SAU-DC-04.scotcourts.local
        Add-ADGroupMember -Identity "GPO SF - Folder Redirection 2" -Members $SAM
        Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        Add-ADGroupMember -Identity $DistributionGroup -Members $SAM 
        Add-ADGroupMember -Identity "DomainShareAccess" -Members $SAM
        Add-ADGroupMember -Identity "All Judicial Studies Access" -Members $SAM
        Add-ADGroupMember -Identity "Judicial Studies Access" -Members $SAM
        Add-ADGroupMember -Identity "GPO SF - Judicial Hub Home Page" -Members $SAM
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "USR-PERS" }
        Set-ADUser -Identity $SAM -Enabled $True 
        Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true
        Set-ADAccountExpiration -Identity $SAM -DateTime $date
        ##Set-CASMailbox -Identity $SAM -DomainController SAU-DC-04.scotcourts.local -PopEnabled $False -OWAEnabled $False -ImapEnabled $False -ActiveSyncEnabled $False
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        Write-host "Username: $SAM  - Password: $pass  - Email address: $mail"
        $objForm.Close() | Out-Null
        MainForm
    }
    Function MainForm {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        [int]$inc = 0
        $JudFirstName = Import-Csv "\\scotcourts.local\data\it\Enterprise Team\UserManagement\Lists\JudicialUsers\FirstNameJud.csv"
        $ManForm = New-Object System.Windows.Forms.Form
        $ManForm.Icon = $Icon
        $ManForm.Size = New-Object System.Drawing.Size(375, 400)
        $ManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $ManForm.MinimizeBox = $False
        $ManForm.MaximizeBox = $False
        $ManForm.FormBorderStyle = 'Fixed3D'
        $Logo = [System.Drawing.Image]::Fromfile('\\scotcourts.local\data\CDiScripts\Scripts\Resources\icons\SCTS.png')
        $pictureBox = New-Object Windows.Forms.PictureBox
        $pictureBox.Width = $Logo.Size.Width
        $pictureBox.Height = $Logo.Size.Height
        $pictureBox.Image = $Logo
        $ManForm.controls.add($pictureBox)
        $ManForm.Text = $WinTitle
        $ManForm.StartPosition = 'CenterScreen'
        $ManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(100, 325)
        $OKButton.Size = New-Object System.Drawing.Size(75, 23)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ManForm.AcceptButton = $OKButton
        $ManForm.Controls.Add($OKButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(175, 325)
        $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $CancelButton.Text = 'Cancel'
        $CancelButton.add_Click( {
                $ManForm.Close()
                $ManForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
        $ManForm.CancelButton = $CancelButton
        $ManForm.Controls.Add($CancelButton)
        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(10, 120)
        $label1.Size = New-Object System.Drawing.Size(280, 20)
        $label1.Text = 'First name:'
        $ManForm.Controls.Add($label1)
        $textBox1 = New-Object System.Windows.Forms.TextBox
        $textBox1.Location = New-Object System.Drawing.Point(10, 140)
        $textBox1.Size = New-Object System.Drawing.Size(330, 20)
        $ManForm.Controls.Add($textBox1)
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10, 165)
        $label2.Size = New-Object System.Drawing.Size(280, 20)
        $label2.Text = 'Last name:'
        $ManForm.Controls.Add($label2)
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10, 210)
        $label3.Size = New-Object System.Drawing.Size(100, 20)
        $label3.Text = 'Account Type:'
        $ManForm.Controls.Add($label3)
        $ContBox = New-Object System.Windows.Forms.CheckBox
        $ContBox.UseVisualStyleBackColor = $True
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 104
        $System_Drawing_Size.Height = 24
        $ContBox.Size = $System_Drawing_Size
        $ContBox.TabIndex = 2
        $ContBox.Text = "Contractor"
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = 160
        $System_Drawing_Point.Y = 210
        $ContBox.Location = $System_Drawing_Point
        $ContBox.DataBindings.DefaultDataSourceUpdateMode = 0
        $ContBox.Name = "Contractor"
        $ManForm.Controls.Add($ContBox)
        $textBox2 = New-Object System.Windows.Forms.TextBox
        $textBox2.Location = New-Object System.Drawing.Point(10, 185)
        $textBox2.Size = New-Object System.Drawing.Size(330, 20)
        $ManForm.Controls.Add($textBox2)
        $BoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
        $BoxRadioButton1.Location = '20,220'
        $BoxRadioButton1.size = '200,40'
        $BoxRadioButton1.Checked = $true 
        $BoxRadioButton1.Text = 'SCTS.'
        $BoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
        $BoxRadioButton2.Location = '20,250'
        $BoxRadioButton2.size = '200,40'
        $BoxRadioButton2.Checked = $false
        $BoxRadioButton2.Text = 'Tribs.'
        $BoxRadioButton3 = New-Object System.Windows.Forms.RadioButton
        $BoxRadioButton3.Location = '20,280'
        $BoxRadioButton3.size = '100,40'
        $BoxRadioButton3.Checked = $false
        $BoxRadioButton3.Text = 'Judicial.'
        $BoxRadioButton3.Add_Click( {
                $JudcomboBox.Enabled = $true
                $BoxRadioButton1.Checked = $false
                $BoxRadioButton2.Checked = $false })
        $BoxRadioButton2.Add_Click( {
                $JudcomboBox.Enabled = $false
                $BoxRadioButton1.Checked = $false
                $BoxRadioButton2.Checked = $true })
        $BoxRadioButton1.Add_Click( {
                $JudcomboBox.Enabled = $false
                $BoxRadioButton1.Checked = $true
                $BoxRadioButton2.Checked = $false })
        $JudcomboBox = New-Object System.Windows.Forms.ComboBox
        $JudcomboBox.Location = '190,290'
        $JudcomboBox.Size = '130,40'
        $JudcomboBox.AutoCompleteMode = 'Suggest'
        $JudcomboBox.AutoCompleteSource = 'ListItems'
        $JudcomboBox.Sorted = $false;
        $JudcomboBox.Enabled = $false;
        $JudcomboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudcomboBox.SelectedItem = $JudcomboBox.Items[0]
        $JudcomboBox.DataSource = $JudFirstName.FirstName 
        $JudcomboBox.add_SelectedIndexChanged( { $JudFirstNameSelect.Text = "$($JudcomboBox.SelectedItem.ToString())" })
        $JudFirstNameSelect = New-Object System.Windows.Forms.Label
        $JudFirstNameSelect.Location = '190,290'
        $JudFirstNameSelect.size = '85,40'
        $JudText = New-Object System.Windows.Forms.Label
        $JudText.Location = '140,290'
        $JudText.size = '50,30'
        $JudText.Text = "Title:" 
        $ManForm.Controls.AddRange(@($JudcomboBox, $JudFirstNameSelect, $JudText))
        $ManForm.Controls.Add($BoxRadioButton3)
        $ManForm.Controls.Add($BoxRadioButton2)
        $ManForm.Controls.Add($BoxRadioButton1)
        $ManForm.Topmost = $true
        $ManForm.Add_Shown( { $textBox1.Select() })
        $result = $ManForm.ShowDialog()
        $firstname = $textBox1.Text
        $lastname = $textBox2.Text
        $Title = $JudcomboBox.Text
        if ($Result -eq 'OK') {
            $charlist1 = [char]97..[char]122
            $charlist2 = [char]65..[char]90
            $charlist3 = [char]48..[char]57
            $charlist4 = [char]33..[char]38 + [char]40..[char]43 + [char]45..[char]46 + [char]64
            $pwdList = @()
            $pwLength = 1 
            For ($i = 0; $i -lt $pwlength; $i++) {
                $pwdList += $charlist1 | Get-Random
                $pwdList += $charlist2 | Get-Random
                $pwdList += $charlist3 | Get-Random
                $pwdList += $charlist4 | Get-Random
                $pwdList += $charlist1 | Get-Random
                $pwdList += $charlist2 | Get-Random
                $pwdList += $charlist3 | Get-Random
                $pwdList += $charlist1 | Get-Random
                $pwdList += $charlist2 | Get-Random
                $pwdList += $charlist3 | Get-Random
            }
            $pass = -join ($pwdList | get-random -count $pwdList.count)
            Write-Host "Please Note: This account will be made with Password: " $pass -ForegroundColor Red
            $password = ConvertTo-SecureString $pass -AsPlainText -Force
            if ($textBox1.Text -eq '') {
                Write-Host "No Firstname entered"
                [System.Windows.Forms.MessageBox]::Show("You need to enter a user name!  Trying to enter blank fields is never a good idea.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            elseif ($textBox2.Text -eq '') {
                Write-Host "No Surname entered"
                [System.Windows.Forms.MessageBox]::Show("You need to enter a user name!  Trying to enter blank fields is never a good idea.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            elseif ($FirstName -match " ") {
                Write-Host "space found in first name"
                [System.Windows.Forms.MessageBox]::Show("Space cannot be in a name", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            elseif ($LastName -match " ") {
                Write-Host "space found in last name"
                [System.Windows.Forms.MessageBox]::Show("Space cannot be in a name", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            Else {
                if ($Result -eq 'OK') {
                    if ($BoxRadioButton1.Checked) {
                        $TempDisplayName = $LastName + ", " + $FirstName
                        $tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower()
                        $DisplayName = $TempDisplayName
                        $samcatch = $tentativeSAM
                        $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                        if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                            do {
                                $inc ++
                                $tentativeSAM = $samcatch + [string]$inc
                                $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                            } 
                            until (-not (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }))
                        }
                        $mail = "$tentativeSAM@ScotCourts.gov.uk"
                        $SAM = $tentativeSAM
                        Write-Host "$TempDisplayName Requested, Sam $SAM & Email $mail"
                        if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                            do {
                                $inc ++
                                $TempDisplayName = $DisplayName + [string]$inc
                            } 
                            until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                        }
                        $DisplayName = $TempDisplayName
                        Write-Host "$DisplayName Set"
                        Write-Host "Sam $SAM & Email $mail Set"
                        if ($ContBox.Checked) {
                            SCTSContForm
                        }
                        else {
                            SCTSForm
                        }
                    }
                    elseif ($BoxRadioButton2.Checked) {
                        $TempDisplayName = $LastName + ", " + $FirstName
                        $tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower()
                        $samcatch = $tentativeSAM
                        $DisplayName = $TempDisplayName
                        $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                        if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                            do {
                                $inc ++
                                $tentativeSAM = $samcatch + [string]$inc
                                $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                            } 
                            until (-not (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }))
                        }
                        $mail = "$tentativeSAM@ScotCourtsTribunals.gov.uk"
                        $SAM = $tentativeSAM
                        Write-Host "$DisplayName Requested, Sam $SAM & Email $mail"
                        if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                            do {
                                $inc ++
                                $TempDisplayName = $DisplayName + [string]$inc
                            } 
                            until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                        }
                        $DisplayName = $TempDisplayName
                        Write-Host "$DisplayName Set"
                        Write-Host "Sam $SAM & Email $mail Set"
                        if ($ContBox.Checked) {
                            TribsContForm
                        }
                        else {
                            Tribsform
                        }
                    }
                    elseif ($BoxRadioButton3.Checked = $True) {
                        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                            if (($Title -eq "Lord") -or ($Title -eq "Lady")) {
                                $TempDisplayName = $Title + " " + $LastName
                                $DisplayName = $TempDisplayName
                                $tentativeSAM = $Title + $LastName
                                $samcatch = $tentativeSAM
                                $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                                if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                                    do {
                                        $inc ++
                                        $tentativeSAM = $samcatch + [string]$inc
                                        $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                                    } 
                                    until (-not (Get-ADUser -Filter { proxyAddresses -eq $tentativeSAM }))
                                }
                                if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                                    do {
                                        $inc ++
                                        $TempDisplayName = $DisplayName + [string]$inc
                                    } 
                                    until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                                }
                                $DisplayName = $TempDisplayName
                            }
                            elseif (($Title -eq "Sheriff") -or ($Title -eq "Sheriffs") -or ($Title -eq "Sheriffp")) {
                                $Sheriff = $firstname.substring(0, 1)
                                $TempDisplayName = $Title + " " + $LastName + " " + $Sheriff
                                $DisplayName = $TempDisplayName
                                $tentativeSAM = $Title + $Sheriff + $LastName
                                $samcatch = $tentativeSAM
                                $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                                if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                                    do {
                                        $inc ++
                                        $tentativeSAM = $samcatch + [string]$inc
                                        $EmailCatch = "smtp:$tentativeSAM@ScotCourts.gov.uk"
                                    } 
                                    until (-not (Get-ADUser -Filter { proxyAddresses -eq $tentativeSAM }))
                                }
                                if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                                    do {
                                        $inc ++
                                        $TempDisplayName = $DisplayName + [string]$inc
                                    } 
                                    until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                                }
                                $DisplayName = $TempDisplayName
                            }
                            else {
                                Write-Host "No title selected"
                                [System.Windows.Forms.MessageBox]::Show("You need to select a title.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                $ManForm.Close()
                                $ManForm.Dispose()
                                break
                            }
                            $SAM = $tentativeSAM
                            Write-Host "$DisplayName Requested, Sam $SAM & Email $mail"
                            if ($ContBox.Checked) {
                                JudContForm
                            }
                            else {
                                Judform
                            }
                        }
                    }
                }
            
            }
        }
    }
    Return MainForm 
}
