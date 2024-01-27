# Create User
# Author        Brian Stark
# Date          23/07/2022      
# Purpose       To allow helpdesk to create accounts correctly
# Requirement   Powershell 7 needed, Must run as Admin.
# 
# Build              
$version = '1.00'
#
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Create\$Date.txt" -append
$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\User.ico'
$Lists = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\Control.csv"
$WinTitle = "Create User v$version."
$DC = "SAU-DC-04.scotcourts.local"
$groups = "GPO SF - Folder Redirection 2", "DomainShareAccess","Wifi-UserCertDeployment"
$Proxies = @()
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos  
Import-PSSession $session
$UserName = $env:username
if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
}
else {
    Function Build {
        $UPN = $SAM + '@scotcourts.gov.uk'
        if ($Grade -eq 'AA/SGB2') {
            $GradeNumber = "1"
        }
        if ($Grade -eq 'AO/SGB1') {
            $GradeNumber = "2"
        }
        if ($Grade -eq 'PS') {
            $GradeNumber = "3"
        }
        if ($Grade -eq 'EO') {
            $GradeNumber = "4"
        }
        if ($Grade -eq 'HEO') {
            $GradeNumber = "5"
        }
        if ($Grade -eq 'SEO') {
            $GradeNumber = "6"
        }
        if ($Grade -eq 'GD7') {
            $GradeNumber = "7"
        }
        if ($Grade -eq 'GD6') {
            $GradeNumber = "8"
        }
        if (($Track -eq "1") -or ($Track -eq "2") -or ($Track -eq "3")) {
            $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
            New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -FirstName $FirstName -LastName $Lastname -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -RemoteRoutingAddress $Routingaddress -DomainController $DC -Password $Password -ResetPasswordOnNextLogon $true
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
            Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true -Office $Office -Enabled $True -Description $Description -Manager $Manager -Department $Department
            foreach ($group in $groups) {
                Add-ADGroupMember -Identity $groups -Members $SAM
            }
            foreach ($Proxy in $Proxies) {
                Set-ADUser -identity $SAM -add @{proxyAddresses = ($Proxy) }
            }
            Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
            Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
            Set-ADUser -Identity $SAM -add @{"extensionattribute5" = $GradeNumber }
        }
        elseif ($Track -eq "4") {
            # RBAC account
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
            New-AdUser -Name $DisplayName -SamAccountName $SAM -GivenName $FirstName -Surname $Lastname -DisplayName $DisplayName -UserPrincipalName $UPN -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server $DC -AccountPassword $Password -passThru
            Start-Sleep -Seconds 5
            Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
            Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
            Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        }
        elseif ($Track -eq "5") {
            # RAS account
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
            New-AdUser -Name $DisplayName -SamAccountName $SAM -GivenName $FirstName -Surname $Lastname -DisplayName $DisplayName -UserPrincipalName $UPN -Office $Office -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server $DC -AccountPassword $Password -passThru
            Start-Sleep -Seconds 5
            Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
            Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
            Add-ADGroupMember -Identity $SecurityGroup -Members $SAM
        }
        elseif ($Track -eq "6") {
            # Clone account
            $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
            New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -FirstName $FirstName -LastName $Lastname -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -RemoteRoutingAddress $Routingaddress -DomainController $DC -Password $Password -ResetPasswordOnNextLogon $true
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
            Set-ADUser -Identity $SAM -ChangePasswordAtLogon $true -Office $Office -Enabled $True -Description $Description -Manager $Manager -Department $Department
            foreach ($group in $groups) {
                Add-ADGroupMember -Identity $groups -Members $SAM
            }
            if ($Attrib2 -eq "USR-PERS") {
                $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
            }
            else {
                $Proxies = @($TribPrimary, $SCSProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
            }
            foreach ($Proxy in $Proxies) {
                Set-ADUser -identity $SAM -add @{proxyAddresses = ($Proxy) }
            }
            Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
            Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
            Set-ADUser -Identity $SAM -add @{"extensionattribute5" = $GradeNumber }
            get-ADuser -identity $UserSamAccountName -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $SAM
        }
        elseif ($Track -eq "7") {
            # Sharedmailbox account
            $Routingaddress = $SAM + '@scotcourtsgovuk.mail.onmicrosoft.com'
            New-RemoteMailbox -Name $DisplayName -SamAccountName $SAM -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU -PrimarySmtpAddress $mail -RemoteRoutingAddress $Routingaddress -DomainController $DC -Shared
            Add-Type -AssemblyName System.Windows.Forms 
            $objForm = New-Object System.Windows.Forms.Form
            $objForm.Text = $WinTitle
            $objForm.Size = New-Object System.Drawing.Size(450, 270)
            $objForm.StartPosition = "CenterScreen"
            $objForm.Controlbox = $false
            $objLabel = New-Object System.Windows.Forms.Label
            $objLabel.Location = New-Object System.Drawing.Size(80, 50) 
            $objLabel.Size = New-Object System.Drawing.Size(300, 120)
            $objLabel.Text = "A New Shared mailbox is being created in AD`nwith the details you entered.`n`nThe New Account will be created in the shared mailbox OU.`n`nPlease Wait."
            $objForm.Controls.Add($objLabel)
            $objForm.Show() | Out-Null
            $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
            Start-Sleep -Seconds 5
            foreach ($group in $groups) {
                Add-ADGroupMember -Identity $groups -Members $SAM
            }
            foreach ($Proxy in $Proxies) {
                Set-ADUser -identity $SAM -add @{proxyAddresses = ($Proxy) }
            }
            Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
            Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
            Set-ADUser -Identity $SAM Office $Office Description $Description
        }
        elseif ($Track -eq "8") {
            # Distlist account
            $pass = $null
            New-DistributionGroup -Name $DisplayName -Type Distribution -OrganizationalUnit $OU -SamAccountName $mail
            Set-adgroup -Identity $User -add @{"extensionattribute2" = $Attrib2 }
            Set-adgroup -Identity $User -add @{"extensionattribute3" = $Ticket }
        }
        elseif ($Track -eq "9") {
            # Securitygroup account
            $pass = $null
            New-DistributionGroup -Name $DisplayName -Type Security -OrganizationalUnit $OU -SamAccountName $mail
            Set-adgroup -Identity $User -add @{"extensionattribute2" = $Attrib2 } 
            Set-adgroup -Identity $User -add @{"extensionattribute3" = $Ticket }
        }
        else {
            Write-host "Tracker error"
        }
        if ($Contractor -eq "1") {
            Set-ADAccountExpiration -Identity $SAM -DateTime $Enddate
            Set-ADUser -Identity $SAM -add @{"extensionattribute4" = "Contractor" }
            Set-ADUser -Identity $SAM 
        }
        $copy = "Username: $SAM  - Password: $pass  - Email address: $mail" | clip
        $objForm.Close() | Out-Null
        MainForm
    }
    Function Selection {
        $Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
        $Managers = get-aduser -Filter { extensionattribute5 -gt '4' } -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties * | Select-Object Displayname | Select-Object -ExpandProperty Displayname
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $SelectionForm = New-Object System.Windows.Forms.Form
        $SelectionForm.width = 800
        $SelectionForm.height = 600
        $SelectionForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $SelectionForm.Controlbox = $false
        $SelectionForm.Icon = $Icon
        $SelectionForm.FormBorderStyle = 'Fixed3D'
        $SelectionForm.Text = $WinTitle
        $SelectionForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $SelectionFormBox1 = New-Object System.Windows.Forms.GroupBox
        $SelectionFormBox1.Location = '20,20'
        $SelectionFormBox1.size = '730,110'
        $SelectionFormBox1.text = '1. The Account Will be created with the following settings:'
        $SelectionFormB1Text1 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text1.Location = '40,40'
        $SelectionFormB1Text1.size = '65,20'
        $SelectionFormB1Text1.Text = 'Name:' 
        $SelectionFormB1Text2 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text2.Location = '140,40'
        $SelectionFormB1Text2.size = '100,20'
        $SelectionFormB1Text2.ForeColor = 'Blue'
        $SelectionFormB1Text2.Text = $DisplayName
        $SelectionFormB1Text4 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text4.Location = '40,80'
        $SelectionFormB1Text4.size = '100,20'
        $SelectionFormB1Text4.Text = 'Email address:'
        $SelectionFormB1Text5 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text5.Location = '140,80'
        $SelectionFormB1Text5.size = '250,20'
        $SelectionFormB1Text5.ForeColor = 'Blue'
        $SelectionFormB1Text5.Text = $mail
        $SelectionFormB1Text6 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text6.Location = '500,40'
        $SelectionFormB1Text6.size = '75,20'
        $SelectionFormB1Text6.Text = 'Logon:'
        $SelectionFormB1Text7 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text7.Location = '575,40'
        $SelectionFormB1Text7.size = '100,20'
        $SelectionFormB1Text7.ForeColor = 'Blue'
        $SelectionFormB1Text7.Text = $tentativeSAM
        $SelectionFormB1Text8 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text8.Location = '500,70'
        $SelectionFormB1Text8.size = '75,20'
        $SelectionFormB1Text8.Text = 'Password:'
        $SelectionFormB1Text9 = New-Object System.Windows.Forms.Label
        $SelectionFormB1Text9.Location = '575,70'
        $SelectionFormB1Text9.size = '120,20'
        $SelectionFormB1Text9.ForeColor = 'Blue'
        $SelectionFormB1Text9.Text = $pass
        $SelectionFormBox1.Controls.AddRange(@($SelectionFormB1Text1, $SelectionFormB1Text2, $SelectionFormB1Text4, $SelectionFormB1Text5, $SelectionFormB1Text6, $SelectionFormB1Text7, $SelectionFormB1Text8, $SelectionFormB1Text9))
        if ($Contractor -eq "1") {
            $Datelabel = New-Object System.Windows.Forms.Label
            $Datelabel.Location = New-Object System.Drawing.Point(20, 480)
            $Datelabel.Size = New-Object System.Drawing.Size(280, 20)
            $Datelabel.Text = "End date: (format '30/12/2022' )"
            $SelectionForm.Controls.Add($Datelabel)
            $DateBox = New-Object System.Windows.Forms.TextBox
            $DateBox.Location = New-Object System.Drawing.Point(20, 510)
            $DateBox.Size = New-Object System.Drawing.Size(330, 20)
            $SelectionForm.Controls.Add($DateBox)
        }
        if ($Track -eq "5") {
            $RasGroups = Get-ADGroup -filter * -searchbase 'OU=RAS,OU=Global,OU=Groups,OU=SCTS,DC=scotcourts,DC=local' -Properties Name | Select-Object name | Select-Object -ExpandProperty name
            $SelectionFormBox2 = New-Object System.Windows.Forms.GroupBox
            $SelectionFormBox2.Location = '20,140'
            $SelectionFormBox2.size = '730,100'
            $SelectionFormBox2.text = "2. Select the RAS Group:"
            $SelectionFormB2Text1 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text1.Location = '20,35'
            $SelectionFormB2Text1.size = '250,30'
            $SelectionFormB2Text1.Text = "Ras Group:" 
            $SelectionFormComboBox1 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormComboBox1.Location = '325,35'
            $SelectionFormComboBox1.Size = '350, 310'
            $SelectionFormComboBox1.AutoCompleteMode = 'Suggest'
            $SelectionFormComboBox1.AutoCompleteSource = 'ListItems'
            $SelectionFormComboBox1.Sorted = $true;
            $SelectionFormComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormComboBox1.DataSource = $RasGroups
            $SelectionFormBox2.Controls.AddRange(@($SelectionFormComboBox1, $SelectionFormB2Text1))
        }
        elseif (($Track -eq "4") -or ($Track -eq "6")) {
            $UserNameList = Get-ADUser -filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
            Write-Output "Pulling User Accounts"
            $SelectionFormBox2 = New-Object System.Windows.Forms.GroupBox
            $SelectionFormBox2.Location = '20,140'
            $SelectionFormBox2.size = '730,120'
            if ($Track -eq "6") {
                $SelectionFormBox2.text = "2. Select the User to Clone:"
            }
            else {
                $SelectionFormBox2.text = "2. Select the User Create RBAC Account for:"
            }
            $SelectionFormB2Text1 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text1.Location = '20,35'
            $SelectionFormB2Text1.size = '250,30'
            $SelectionFormB2Text1.Text = "Existing User:" 
            $SelectionFormComboBox1 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormComboBox1.Location = '325,35'
            $SelectionFormComboBox1.Size = '350, 310'
            $SelectionFormComboBox1.AutoCompleteMode = 'Suggest'
            $SelectionFormComboBox1.AutoCompleteSource = 'ListItems'
            $SelectionFormComboBox1.Sorted = $true;
            $SelectionFormComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormComboBox1.DataSource = $UsernameList
            if ($Track -eq "4") {
                $RBACGroups = Get-ADGroup -filter * -searchbase 'OU=RBAC,OU=Groups,OU=SCTS,DC=scotcourts,DC=local' -Properties Name | Select-Object name | Select-Object -ExpandProperty name
                $SelectionFormComboBox1.add_SelectedIndexChanged( { $SelectionFormB1Text2.Text = "$($SelectionFormComboBox1.SelectedItem.ToString() + " - RBAC")" })
                $SelectionFormComboBox1.add_SelectedIndexChanged( { $SelectionFormB1Text7.Text = "$((Get-ADUser -filter { DisplayName -eq $SelectionFormComboBox1.Text } -Properties * | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName) + "_A")" })
                $SelectionFormB2Text2 = New-Object System.Windows.Forms.Label
                $SelectionFormB2Text2.Location = '20,75'
                $SelectionFormB2Text2.size = '100,30'
                $SelectionFormB2Text2.Text = "RBAC Group:" 
                $SelectionFormB2comboBox2 = New-Object System.Windows.Forms.ComboBox
                $SelectionFormB2comboBox2.Location = '325,75'
                $SelectionFormB2comboBox2.Size = '350,40'
                $SelectionFormB2comboBox2.AutoCompleteMode = 'Suggest'
                $SelectionFormB2comboBox2.AutoCompleteSource = 'ListItems'
                $SelectionFormB2comboBox2.Sorted = $false;
                $SelectionFormB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
                $SelectionFormB2comboBox2.DataSource = $RBACGroups
                $SelectionFormB2comboBox2.SelectedItem = $SelectionFormB2comboBox2.Items[0]
                $SelectionFormB2comboBox2.add_SelectedIndexChanged( { $SelectionFormB2comboBoxDescriptionSelect.Text = "$($SelectionFormB2comboBox2.SelectedItem.ToString())" })
                $SelectionFormBox2.Controls.AddRange(@($SelectionFormB2Text2, $SelectionFormB2comboBox2))
            }
            $SelectionFormBox2.Controls.AddRange(@($SelectionFormComboBox1, $SelectionFormB2Text1))
        }
        elseif ($Track -eq "7") {
            $SelectionFormBox2 = New-Object System.Windows.Forms.GroupBox
            $SelectionFormBox2.Location = '20,140'
            $SelectionFormBox2.size = '730,300'
            $SelectionFormBox2.text = "2. Select the new users AD details:"
            $SelectionFormB2Text1 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text1.Location = '20,40'
            $SelectionFormB2Text1.size = '100,30'
            $SelectionFormB2Text1.Text = "Owner:" 
            $SelectionFormComboBox1 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormComboBox1.Location = '325,35'
            $SelectionFormComboBox1.Size = '350, 310'
            $SelectionFormComboBox1.AutoCompleteMode = 'Suggest'
            $SelectionFormComboBox1.AutoCompleteSource = 'ListItems'
            $SelectionFormComboBox1.Sorted = $true;
            $SelectionFormComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormComboBox1.DataSource = $UsernameList
            $SelectionFormB2Text2 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text2.Location = '20,75'
            $SelectionFormB2Text2.size = '100,30'
            $SelectionFormB2Text2.Text = "Description:" 
            $DesctextBox = New-Object System.Windows.Forms.TextBox
            $DesctextBox.Location = New-Object System.Drawing.Point(325, 75)
            $DesctextBox.Size = New-Object System.Drawing.Size(330, 20)
            $SCTSGroupButton = New-Object System.Windows.Forms.RadioButton
            $SCTSGroupButton.Location = '20,100'
            $SCTSGroupButton.size = '200,30'
            $SCTSGroupButton.Checked = $true 
            $SCTSGroupButton.Text = 'SCTS.'
            $TribsGroupButton = New-Object System.Windows.Forms.RadioButton
            $TribsGroupButton.Location = '20,150'
            $TribsGroupButton.size = '200,30'
            $TribsGroupButton.Checked = $false
            $TribsGroupButton.Text = 'Tribs.'
            $SelectionFormBox2.Controls.AddRange(@($SelectionFormB2Text2, $SelectionFormB2Text1, $SelectionFormComboBox1, $SelectionFormB2Text2, $DesctextBox, $TribsGroupButton, $SCTSGroupButton))
        }
        else {
            $SelectionFormBox2 = New-Object System.Windows.Forms.GroupBox
            $SelectionFormBox2.Location = '20,140'
            $SelectionFormBox2.size = '730,300'
            $SelectionFormBox2.text = "2. Select the new users AD details:"
            $SelectionFormB2Text1 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text1.Location = '20,40'
            $SelectionFormB2Text1.size = '350,30'
            $SelectionFormB2Text1.Text = "Office field in AD:" 
            $SelectionFormB2Text2 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text2.Location = '20,75'
            $SelectionFormB2Text2.size = '350,30'
            $SelectionFormB2Text2.Text = "Description field in AD:" 
            $SelectionFormB2Text3 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text3.Location = '20,105'
            $SelectionFormB2Text3.size = '350,30'
            $SelectionFormB2Text3.Text = "Distribution List field in AD:" 
            $SelectionFormB2Text4 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text4.Location = '20,140'
            $SelectionFormB2Text4.size = '370,30'
            $SelectionFormB2Text4.Text = "Security Group field in AD (to access N drives):" 
            $SelectionFormB2comboBox1 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox1.Location = '425,40'
            $SelectionFormB2comboBox1.Size = '250,40'
            $SelectionFormB2comboBox1.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox1.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox1.Sorted = $false;
            $SelectionFormB2comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox1.SelectedItem = $SelectionFormB2comboBox1.Items[0]
            $SelectionFormB2comboBox1.add_SelectedIndexChanged( { $SelectionFormB2comboBoxOfficeSelect.Text = "$($SelectionFormB2comboBox1.SelectedItem.ToString())" })
            $SelectionFormB2comboBox2 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox2.Location = '425,75'
            $SelectionFormB2comboBox2.Size = '250,40'
            $SelectionFormB2comboBox2.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox2.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox2.Sorted = $false;
            $SelectionFormB2comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox2.SelectedItem = $SelectionFormB2comboBox2.Items[0]
            $SelectionFormB2comboBox2.add_SelectedIndexChanged( { $SelectionFormB2comboBoxDescriptionSelect.Text = "$($SelectionFormB2comboBox2.SelectedItem.ToString())" })
            $SelectionFormB2comboBox3 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox3.Location = '425,105'
            $SelectionFormB2comboBox3.Size = '250,40'
            $SelectionFormB2comboBox3.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox3.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox3.Sorted = $false;
            $SelectionFormB2comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox3.SelectedItem = $SelectionFormB2comboBox3.Items[0]
            $SelectionFormB2comboBox3.add_SelectedIndexChanged( { $SelectionFormB2comboBoxDistributionSelect.Text = "$($SelectionFormB2comboBox3.SelectedItem.ToString())" })
            $SelectionFormB2comboBox4 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox4.Location = '425,140'
            $SelectionFormB2comboBox4.Size = '250,40'
            $SelectionFormB2comboBox4.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox4.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox4.Sorted = $false;
            $SelectionFormB2comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox4.SelectedItem = $SelectionFormB2comboBox4.Items[0]
            $SelectionFormB2comboBox4.add_SelectedIndexChanged( { $SelectionFormB2comboBoxSecuritySelect.Text = "$($SelectionFormB2comboBox4.SelectedItem.ToString())" })
            $SelectionFormB2comboBoxOfficeSelect = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxOfficeSelect.Location = '20,600'
            $SelectionFormB2comboBoxOfficeSelect.size = '350,40'
            $SelectionFormB2comboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxDescriptionSelect.Location = '20,650'
            $SelectionFormB2comboBoxDescriptionSelect.size = '350,40'
            $SelectionFormB2comboBoxDistributionSelect = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxDistributionSelect.Location = '20,700'
            $SelectionFormB2comboBoxDistributionSelect.size = '350,40'
            $SelectionFormB2comboBoxSecuritySelect = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxSecuritySelect.Location = '20,750'
            $SelectionFormB2comboBoxSecuritySelect.size = '350,40'
            $SelectionFormB2Text5 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text5.Location = '20,170'
            $SelectionFormB2Text5.size = '370,30'
            $SelectionFormB2Text5.Text = "Grade:" 
            $SelectionFormB2comboBox5 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox5.Location = '425,170'
            $SelectionFormB2comboBox5.Size = '250,40'
            $SelectionFormB2comboBox5.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox5.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox5.Sorted = $false;
            $SelectionFormB2comboBox5.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox5.SelectedItem = $SelectionFormB2comboBox5.Items[0]
            $SelectionFormB2comboBox5.add_SelectedIndexChanged( { $SelectionFormB2comboBoxGrade.Text = "$($SelectionFormB2comboBox5.SelectedItem.ToString())" })
            $SelectionFormB2comboBoxGrade = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxGrade.Location = '20,800'
            $SelectionFormB2comboBoxGrade.size = '350,40'
            $SelectionFormB2comboBox5.DataSource = $Lists.Grade | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            $SelectionFormB2Text6 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text6.Location = '20,200'
            $SelectionFormB2Text6.size = '370,30'
            $SelectionFormB2Text6.Text = "Department:" 
            $SelectionFormB2comboBox6 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox6.Location = '425,200'
            $SelectionFormB2comboBox6.Size = '250,40'
            $SelectionFormB2comboBox6.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox6.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox6.Sorted = $false;
            $SelectionFormB2comboBox6.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox6.SelectedItem = $SelectionFormB2comboBox6.Items[0]
            $SelectionFormB2comboBox6.add_SelectedIndexChanged( { $SelectionFormB2comboBoxDepartment.Text = "$($SelectionFormB2comboBox6.SelectedItem.ToString())" })
            $SelectionFormB2comboBoxDepartment = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxDepartment.Location = '20,850'
            $SelectionFormB2comboBoxDepartment.size = '350,40'
            $SelectionFormB2comboBox6.DataSource = $Lists.Department | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            $SelectionFormB2Text7 = New-Object System.Windows.Forms.Label
            $SelectionFormB2Text7.Location = '20,230'
            $SelectionFormB2Text7.size = '370,30'
            $SelectionFormB2Text7.Text = "Manager:" 
            $SelectionFormB2comboBox7 = New-Object System.Windows.Forms.ComboBox
            $SelectionFormB2comboBox7.Location = '425,230'
            $SelectionFormB2comboBox7.Size = '250,40'
            $SelectionFormB2comboBox7.AutoCompleteMode = 'Suggest'
            $SelectionFormB2comboBox7.AutoCompleteSource = 'ListItems'
            $SelectionFormB2comboBox7.Sorted = $false;
            $SelectionFormB2comboBox7.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
            $SelectionFormB2comboBox7.SelectedItem = $SelectionFormB2comboBox6.Items[0]
            $SelectionFormB2comboBox7.add_SelectedIndexChanged( { $SelectionFormB2comboBoxManager.Text = "$($SelectionFormB2comboBox7.SelectedItem.ToString())" })
            $SelectionFormB2comboBoxManager = New-Object System.Windows.Forms.Label
            $SelectionFormB2comboBoxManager.Location = '20,900'
            $SelectionFormB2comboBoxManager.size = '350,40'
            $SelectionFormB2comboBox7.DataSource = $Managers
            if ($Track -eq "1") {
                $SelectionFormB2comboBox4.DataSource = $Lists.Description | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox3.DataSource = $Distributionlists 
                $SelectionFormB2comboBox2.DataSource = $Lists.Description | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox1.DataSource = $Lists.OfficeList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            }
            if ($Track -eq "2") {
                $SelectionFormB2comboBox4.DataSource = $Lists.DescriptionTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox3.DataSource = $Lists.DistributionTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox2.DataSource = $Lists.DescriptionTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox1.DataSource = $Lists.OfficeTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            }
            if ($Track -eq "3") {
                $SelectionFormB2comboBox4.DataSource = $Lists.descriptionJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox3.DataSource = $Lists.distributionJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox2.DataSource = $Lists.descriptionJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormB2comboBox1.DataSource = $Lists.officeJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            }
                
            $SelectionFormBox2.Controls.AddRange(@($SelectionFormB2comboBoxSecuritySelect, $SelectionFormB2comboBoxDistributionSelect, $SelectionFormB2comboBoxDescriptionSelect, $SelectionFormB2comboBoxOfficeSelect, $SelectionFormB2comboBox4, $SelectionFormB2comboBox3, $SelectionFormB2comboBox2, $SelectionFormB2comboBox1, $SelectionFormB2Text4, $SelectionFormB2Text3, $SelectionFormB2Text2, $SelectionFormB2Text1, $SelectionFormB2comboBoxGrade, $SelectionFormB2comboBox5, $SelectionFormB2Text5, $SelectionFormB2Text6, $SelectionFormB2comboBox6, $SelectionFormB2comboBoxDepartment, $SelectionFormB2Text7, $SelectionFormB2comboBox7, $SelectionFormB2comboBoxManager))
        }
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '525,500'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '625,500'
        $CancelButton.Size = '100,40'
        $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
        $CancelButton.Text = 'Cancel back to MainForm'
        $CancelButton.add_Click( {
                $SelectionForm.Close()
                $SelectionForm.Dispose()
                Return MainForm })
        $SelectionForm.Controls.AddRange(@($SelectionFormBox1, $SelectionFormBox2, $OKButton, $CancelButton))
        $SelectionForm.AcceptButton = $OKButton
        $SelectionForm.CancelButton = $CancelButton
        $SelectionForm.Add_Shown( { $SelectionForm.Activate() })    
        $dialogResult = $SelectionForm.ShowDialog()
        if ($dialogResult -eq 'OK') {
            Write-Host $Track
            if ($Contractor -eq "1") {
                $Enddate = $DateBox.Text
            }
            if (($Track -eq "1") -or ($Track -eq "2") -or ($Track -eq "3")) {
                $Office = $SelectionFormB2comboBoxOfficeSelect.text
                $Description = $SelectionFormB2comboBoxDescriptionSelect.Text
                $DistributionGroup = $SelectionFormB2comboBoxDistributionSelect.Text
                $SecurityGroup = $SelectionFormB2comboBoxSecuritySelect.Text   
                $Grade = $SelectionFormB2comboBoxGrade.Text
                $Department = $SelectionFormB2comboBoxDepartment.Text
                $Manager = Get-aduser -Identity $SelectionFormB2comboBoxManager.Text -Properties * | Select-Object DistinguishedName | Select-Object -ExpandProperty DistinguishedName 
                $groups += $Office, $SecurityGroup, $DistributionGroup
                if (($Track -eq "1") -or ($Track -eq "3")) {
                    $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                }
                else {
                    $Proxies = @($TribPrimary, $SCSProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                }
                $SelectionForm.Close()
                $SelectionForm.Dispose()
                Build
            }
            if ($Track -eq "4") {
                $Sam = $SelectionFormB1Text7.Text
                $displayname = $SelectionFormB1Text2.Text
                $SecurityGroup = $SelectionFormB2comboBoxDescriptionSelect.Text
                $OU = "OU=RBAC Admins,OU=User Accounts (Admin),OU=SCTS,DC=scotcourts,DC=local"
                $SelectionForm.Close()
                $SelectionForm.Dispose()
                Build
            }
            if ($Track -eq "5") {
                $SecurityGroup = $SelectionFormB2comboBoxDescriptionSelect.Text
                $OU = "OU=External Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
                $SelectionForm.Close()
                $SelectionForm.Dispose()
                Build
            }
            if ($Track -eq "6") {
                $UserSamAccountName = Get-ADUser -filter { DisplayName -eq $CloneformComboBox1.Text } -Properties * | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName 
                Write-Host "$UserSamAccountName selected for Cloning."
                if (($UserSamAccountName | Measure-Object).count -ne 1) { MainForm }
                $Tester = get-aduser $User -properties * | Select-object Office
                if (($Tester -in $Lists.OfficeList) -or ($Tester -in $Lists.OfficeTribs) -or ($Tester -in $Lists.officeJud)) {
                    $Type = get-ADuser -identity $UserSamAccountName -properties * | Select-Object extensionattribute2 | Select-Object -ExpandProperty extensionattribute2
                    if ($Type = "USR-PERS") {
                        Write-Host "$UserSamAccountName is SCS Staff."
                        $Office = get-ADuser -identity $UserSamAccountName -properties * | Select-Object Office | Select-Object -ExpandProperty Office
                        $Description = get-ADuser -identity $UserSamAccountName -properties * | Select-Object Description | Select-Object -ExpandProperty Description
                        $Attrib2 = "USR-PERS"
                        $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                    }
                    else {
                        Write-Host "$UserSamAccountName Is Tribs Staff."
                        $Office = get-ADuser -identity $UserSamAccountName -properties * | Select-Object Office | Select-Object -ExpandProperty Office
                        $Description = get-ADuser -identity $UserSamAccountName -properties * | Select-Object Description | Select-Object -ExpandProperty Description
                        $Attrib2 = "USR-PERT"
                        $Proxies = @($TribPrimary, $SCSProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                    }
                    $SelectionForm.Close()
                    $SelectionForm.Dispose()
                    Build
                }
                else {
                    $SelectionForm.Close()
                    $SelectionForm.Dispose()
                    MainForm
                }
            }
            if ($Track -eq "7") {
                #Shared Mailbox
                $Owner = Get-ADUser -filter { DisplayName -eq $OwnerComboBox1.Text } -Properties * | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName
                $Office = "Owner: $Owner" 
                $Description = $DesctextBox.Text
                if ($SCTSGroupButton.Checked) {
                    $Proxies = @($SCSPrimary, $TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                    $Attrib2 = "MBS-SHAR"
                }
                else {
                    $Proxies = @($TribPrimary, $SCSProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4)
                    $Attrib2 = "MBT-SHAR"
                }
                $OU = "OU=Shared Mailboxes,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local"
                $SelectionForm.Close()
                $SelectionForm.Dispose()
                Build
            }
            if ($Track -eq "8") {

            }
            if ($Track -eq "9") {

            }
        }
    }

    Function MainForm {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        [int]$incA = 0
        [int]$incB = 0
        $OU = "OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
        $Track = "0"
        $Contractor = "0"
        $ManForm = New-Object System.Windows.Forms.Form
        $ManForm.Icon = $Icon
        $ManForm.Size = New-Object System.Drawing.Size(375, 575)
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
        $OKButton.Location = New-Object System.Drawing.Point(175, 475)
        $OKButton.Size = New-Object System.Drawing.Size(75, 23)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ManForm.AcceptButton = $OKButton
        $ManForm.Controls.Add($OKButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(250, 475)
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
        $textBox2 = New-Object System.Windows.Forms.TextBox
        $textBox2.Location = New-Object System.Drawing.Point(10, 190)
        $textBox2.Size = New-Object System.Drawing.Size(330, 20)
        $ManForm.Controls.Add($textBox2)
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10, 215)
        $label4.Size = New-Object System.Drawing.Size(280, 20)
        $label4.Text = 'Ticket:'
        $ManForm.Controls.Add($label4)
        $textBox3 = New-Object System.Windows.Forms.TextBox
        $textBox3.Location = New-Object System.Drawing.Point(10, 235)
        $textBox3.Size = New-Object System.Drawing.Size(330, 20)
        $ManForm.Controls.Add($textBox3)
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10, 270)
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
        $System_Drawing_Point.Y = 270
        $ContBox.Location = $System_Drawing_Point
        $ContBox.DataBindings.DefaultDataSourceUpdateMode = 0
        $ContBox.Name = "Contractor"
        $ManForm.Controls.Add($ContBox)
        $SCTSButton = New-Object System.Windows.Forms.RadioButton
        $SCTSButton.Location = '20,300'
        $SCTSButton.size = '200,30'
        $SCTSButton.Checked = $true 
        $SCTSButton.Text = 'SCTS.'
        $TribsButton = New-Object System.Windows.Forms.RadioButton
        $TribsButton.Location = '20,330'
        $TribsButton.size = '200,30'
        $TribsButton.Checked = $false
        $TribsButton.Text = 'Tribs.'
        $JudButton = New-Object System.Windows.Forms.RadioButton
        $JudButton.Location = '20,360'
        $JudButton.size = '100,30'
        $JudButton.Checked = $false
        $JudButton.Text = 'Judicial.'
        $CloneButton = New-Object System.Windows.Forms.RadioButton
        $CloneButton.Location = '20,390'
        $CloneButton.size = '100,30'
        $CloneButton.Checked = $false
        $CloneButton.Text = 'Clone.'
        $RasButton = New-Object System.Windows.Forms.RadioButton
        $RasButton.Location = '20,420'
        $RasButton.size = '100,30'
        $RasButton.Checked = $false
        $RasButton.Text = 'RAS.'
        $RBACButton = New-Object System.Windows.Forms.RadioButton
        $RBACButton.Location = '20,450'
        $RBACButton.size = '100,30'
        $RBACButton.Checked = $false
        $RBACButton.Text = 'RBAC.'
        $SMBButton = New-Object System.Windows.Forms.RadioButton
        $SMBButton.Location = '20,480'
        $SMBButton.size = '100,30'
        $SMBButton.Checked = $false
        $SMBButton.Text = 'Shared.'
        $CloneButton.Add_Click( {
                $CloneButton.Checked = $True
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $false
                $TribsButton.Checked = $false
                $RasButton.Checked = $false
                $JudButton.Checked = $false
                $SMBButton.Checked = $false 
                $textBox1.Enabled = $True
                $textBox2.Enabled = $True })
        $JudButton.Add_Click( {
                $JudcomboBox.Enabled = $true
                $JudButton.Checked = $true
                $SCTSButton.Checked = $false
                $RasButton.Checked = $false
                $CloneButton.Checked = $false
                $TribsButton.Checked = $false
                $SMBButton.Checked = $false  
                $textBox1.Enabled = $True
                $textBox2.Enabled = $True })
        $TribsButton.Add_Click( {
                $JudButton.Checked = $false
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $false
                $RasButton.Checked = $false
                $CloneButton.Checked = $false
                $TribsButton.Checked = $true
                $SMBButton.Checked = $false  
                $textBox1.Enabled = $True
                $textBox2.Enabled = $True })
        $SCTSButton.Add_Click( {
                $JudButton.Checked = $false
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $true
                $RasButton.Checked = $false
                $CloneButton.Checked = $false
                $TribsButton.Checked = $false
                $SMBButton.Checked = $false  
                $textBox1.Enabled = $True
                $textBox2.Enabled = $True })
        $RasButton.Add_Click( {
                $RasButton.Checked = $True
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $false
                $TribsButton.Checked = $false
                $JudButton.Checked = $false
                $CloneButton.Checked = $false
                $SMBButton.Checked = $false 
                $textBox1.Enabled = $True
                $textBox2.Enabled = $True })
        $RBACButton.Add_Click( {
                $RBACButton.Checked = $True
                $RasButton.Checked = $false
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $false
                $TribsButton.Checked = $false
                $JudButton.Checked = $false
                $CloneButton.Checked = $false
                $SMBButton.Checked = $false 
                $textBox1.Enabled = $false
                $textBox2.Enabled = $false })
        $SMBButton.Add_Click( {
                $SMBButton.Checked = $True
                $RasButton.Checked = $false
                $JudcomboBox.Enabled = $false
                $SCTSButton.Checked = $false
                $TribsButton.Checked = $false
                $JudButton.Checked = $false
                $CloneButton.Checked = $false
                $RBACButton.Checked = $false 
                $textBox1.Enabled = $True
                $textBox2.Enabled = $false })
        $JudcomboBox = New-Object System.Windows.Forms.ComboBox
        $JudcomboBox.Location = '190,360'
        $JudcomboBox.Size = '130,40'
        $JudcomboBox.AutoCompleteMode = 'Suggest'
        $JudcomboBox.AutoCompleteSource = 'ListItems'
        $JudcomboBox.Sorted = $false;
        $JudcomboBox.Enabled = $false;
        $JudcomboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $JudcomboBox.SelectedItem = $JudcomboBox.Items[0]
        $JudcomboBox.DataSource = $Lists.FirstNameJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $JudcomboBox.add_SelectedIndexChanged( { $JudFirstNameSelect.Text = "$($JudcomboBox.SelectedItem.ToString())" })
        $JudFirstNameSelect = New-Object System.Windows.Forms.Label
        $JudFirstNameSelect.Location = '190,360'
        $JudFirstNameSelect.size = '85,40'
        $JudText = New-Object System.Windows.Forms.Label
        $JudText.Location = '140,360'
        $JudText.size = '50,30'
        $JudText.Text = "Title:" 
        $ManForm.Controls.AddRange(@($JudcomboBox, $JudFirstNameSelect, $JudText))
        $ManForm.Controls.Add($RBACButton)
        $ManForm.Controls.Add($RasButton)
        $ManForm.Controls.Add($CloneButton)
        $ManForm.Controls.Add($JudButton)
        $ManForm.Controls.Add($TribsButton)
        $ManForm.Controls.Add($SCTSButton)
        $ManForm.Controls.Add($SMBButton)
        $ManForm.Topmost = $true
        $ManForm.Add_Shown( { $textBox1.Select() })
        $result = $ManForm.ShowDialog()
        $firstname = $textBox1.Text
        $lastname = $textBox2.Text
        $Title = $JudcomboBox.Text
        $TicketTrace = $textBox3.Text
        $charlist1 = [char]97..[char]122
        $charlist2 = [char]65..[char]90
        $charlist3 = [char]48..[char]57
        $charlist4 = [char]33..[char]38 + [char]40..[char]43 + [char]45..[char]46 + [char]64
        $pwdList = @()
        $pwLength = 2
        For ($i = 0; $i -lt $pwlength; $i++) {
            $pwdList += $charlist1 | Get-Random
            $pwdList += $charlist2 | Get-Random
            $pwdList += $charlist3 | Get-Random
            $pwdList += $charlist4 | Get-Random
            $pwdList += $charlist1 | Get-Random
            $pwdList += $charlist2 | Get-Random
            $pwdList += $charlist3 | Get-Random

        }
        $pass = -join ($pwdList | get-random -count $pwdList.count)
        Write-Host "Please Note: This account will be made with Password: " $pass -ForegroundColor Red
        $password = ConvertTo-SecureString $pass -AsPlainText -Force
        if ($Result -eq 'OK') {
            if ($RBACButton.Checked) {
                $Attrib2 = "ADMIN-USER"
                $DisplayName = "Select User"
                $mail = "None"
                $Track = "4"
                $OU = "OU=RBAC Admins,OU=User Accounts (Admin),OU=SCTS,DC=scotcourts,DC=local"
                Write-Host "RBAC"
                Selection
            }
            if (($textBox1.Text -eq '') -or ($FirstName -match " ") -or ($LastName -match " ")) {
                Write-output "error in input"
                [System.Windows.Forms.MessageBox]::Show("You need to enter a user name!  blank entries or spaces are not aloowed.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            } 
            $Ticket = "Created on Ticket : $TicketTrace"
            if ($RasButton.Checked) {
                Write-Host "RasButton"
                $TempDisplayName = $LastName + ", " + $FirstName + " - RAS"
                $tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower() + "RAS"
                $DisplayName = $TempDisplayName 
                $samcatch = $tentativeSAM
                $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                    do {
                        $incA ++
                        $tentativeSAM = $samcatch + [string]$incA
                        $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                    } 
                    until (-not (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }))
                }
                if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                    do {
                        $incB ++
                        $TempDisplayName = $DisplayName + [string]$incB
                    } 
                    until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                }
                $DisplayName = $TempDisplayName
                $mail = "$tentativeSAM@scotcourts.gov.uk"
                $SAM = $tentativeSAM
                Write-Host "RAS - $TempDisplayName Requested, Sam $SAM"
                $Groups = "$null"
                $Attrib2 = "USR-RAS"
                $OU = "OU=External Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
                $Track = "5"
            }
            elseif ($SMBButton.Checked) {
                $Track = "7"
                Write-Host "Shared"
                $tentativeSAM = $FirstName
                
            }
            elseif ($JudButton.Checked) {
                $Groups += "All Judicial Studies Access", "Judicial Studies Access", "GPO SF - Judicial Hub Home Page"
                $Attrib2 = "USR-PERJ"
                $Track = "3"
                if (($Title -eq "Lord") -or ($Title -eq "Lady")) {
                    $TempDisplayName = $Title + " " + $LastName
                    $DisplayName = $TempDisplayName
                    $tentativeSAM = $Title + $LastName
                    $samcatch = $tentativeSAM
                    $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                    if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                        do {
                            $incA ++
                            $tentativeSAM = $samcatch + [string]$incA
                            $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                        } 
                        until (-not (Get-ADUser -Filter { proxyAddresses -eq $tentativeSAM }))
                    }
                    if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                        do {
                            $incB ++
                            $TempDisplayName = $DisplayName + [string]$incB
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
                    $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                    if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                        do {
                            $incA ++
                            $tentativeSAM = $samcatch + [string]$incA
                            $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                        } 
                        until (-not (Get-ADUser -Filter { proxyAddresses -eq $tentativeSAM }))
                    }
                    if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                        do {
                            $incB ++
                            $TempDisplayName = $DisplayName + [string]$incB
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
                    Pause
                }
                $SAM = $tentativeSAM
                $Attrib2 = "USR-PERJ"
                Write-Host "Judicial - $DisplayName Requested, Sam $SAM & Email $mail"
            }
            else {
                Write-host "SCTS or Tribs"
                $TempDisplayName = $LastName + ", " + $FirstName
                $tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower()
                $DisplayName = $TempDisplayName 
                $samcatch = $tentativeSAM
                $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                if (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }) {    
                    do {
                        $incA ++
                        $tentativeSAM = $samcatch + [string]$incA
                        $EmailCatch = "smtp:$tentativeSAM@scotcourts.gov.uk"
                    } 
                    until (-not (Get-ADUser -Filter { proxyAddresses -eq $EmailCatch }))
                }
                $mail = "$tentativeSAM@scotcourts.gov.uk"
                $SAM = $tentativeSAM
                Write-Host "$TempDisplayName Requested, Sam $SAM & Email $mail"
                if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                    do {
                        $incB ++
                        $TempDisplayName = $DisplayName + [string]$incB
                    } 
                    until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
                }
                $DisplayName = $TempDisplayName
                Write-Host "$DisplayName Set"
                Write-Host "Sam $SAM & Email $mail Set"
                if ($TribsButton.Checked) {
                    $Groups += "All Users Tribunals", "acl_All STS Users_readwrite", "acl_All_Tribunals_Users"
                    $Attrib2 = "USR-PERT"
                    $Track = "2"
                    Write-Host "Tribs"
                }
                elseif ($CloneButton.Checked) {
                    $Track = "6"
                    Write-Host "Clone"
                }
                
                else {
                    $Attrib2 = "USR-PERS"
                    $Track = "1"
                    Write-Host "SCTS"
                }
            }
            if ($ContBox.Checked) {
                $Contractor = "1"
            }
            else {
                $Contractor = "0"
            }
            $TribProxy = "smtp:$SAM@scotcourtstribunals.gov.uk"
            $SCSPrimary = "SMTP:$SAM@scotcourts.gov.uk"
            $TribPrimary = "SMTP:$SAM@scotcourtstribunals.gov.uk"
            $SCSProxy = "smtp:$SAM@scotcourts.gov.uk"
            $newProxy1 = "smtp:$SAM@scotcourts.pnn.gov.uk"
            $newProxy2 = "smtp:$SAM@scotcourtstribunals.pnn.gov.uk"
            $newProxy3 = "smtp:$SAM@scotcourtsgovuk.mail.onmicrosoft.com"
            $newProxy4 = "X400:C=GB;A=CWMAIL;P=SCS;O=SCOTTISH COURTS;S=" + $Lastname + ";G=" + $FirstName + ";"
            Selection
        }
    }
    Return MainForm
}