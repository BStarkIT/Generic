$version = '1.00'
#
$Date = Get-Date -Format "dd-MM-yyyy"
#Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Change\$Date.txt" -append
$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\User.ico'
$Lists = Import-Csv "\\scotcourts.local\data\CDiScripts\Scripts\Resources\Lists\Control.csv"
$WinTitle = "Create User v$version."
$OU = "OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local"
$DC = "SAU-DC-04.scotcourts.local"
$groups = "GPO SF - Folder Redirection 2", "DomainShareAccess"
$Proxies = @()
#$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://SAU-EXCHANGE-01.scotcourts.local/powershell -Authentication Kerberos  
#Import-PSSession $session
$UserName = $env:username
$UserNameList = Get-ADUser -filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
<#if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
}
else {
#>

function change {
    #$Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name
    $Managers = get-aduser -Filter { extensionattribute5 -gt '3' } -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties * | Select-Object -ExpandProperty Displayname
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $changeForm = New-Object System.Windows.Forms.Form
    $changeForm.width = 800
    $changeForm.height = 600
    $changeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $changeForm.Controlbox = $false
    $changeForm.Icon = $Icon
    $changeForm.FormBorderStyle = 'Fixed3D'
    $changeForm.Text = $WinTitle
    $changeForm.Font = New-Object System.Drawing.Font('Ariel', 10)
    $changeFormBox1 = New-Object System.Windows.Forms.GroupBox
    $changeFormBox1.Location = '20,20'
    $changeFormBox1.size = '730,110'
    $changeFormBox1.text = '1. User to be changed:'
    $changeFormB1Text1 = New-Object System.Windows.Forms.Label
    $changeFormB1Text1.Location = '40,40'
    $changeFormB1Text1.size = '65,20'
    $changeFormB1Text1.Text = 'Name:' 
    $changeFormB1Text2 = New-Object System.Windows.Forms.Label
    $changeFormB1Text2.Location = '140,40'
    $changeFormB1Text2.size = '200,20'
    $changeFormB1Text2.ForeColor = 'Blue'
    $changeFormB1Text2.Text = $DisplayName
    $changeFormBox1.Controls.AddRange(@($changeFormB1Text1, $changeFormB1Text2))
    if ($Track -eq "5") {
        # Contractor
        $TCEnddate = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty AccountExpirationDate
        $CEnddate = $TCEnddate.ToString('dd MMM yyyy')
        $changeFormB1Text3 = New-Object System.Windows.Forms.Label
        $changeFormB1Text3.Location = '40,60'
        $changeFormB1Text3.size = '65,20'
        $changeFormB1Text3.Text = 'end date:' 
        $changeFormB1Text4 = New-Object System.Windows.Forms.Label
        $changeFormB1Text4.Location = '140,60'
        $changeFormB1Text4.size = '200,20'
        $changeFormB1Text4.ForeColor = 'Blue'
        $changeFormB1Text4.Text = $CEnddate
        $changeFormBox1.Controls.AddRange(@($changeFormB1Text3, $changeFormB1Text4))
        $Datelabel = New-Object System.Windows.Forms.Label
        $Datelabel.Location = New-Object System.Drawing.Point(20, 480)
        $Datelabel.Size = New-Object System.Drawing.Size(280, 20)
        $Datelabel.Text = "End date: (format '30/12/2022' )"
        $changeForm.Controls.Add($Datelabel)
        $DateBox = New-Object System.Windows.Forms.TextBox
        $DateBox.Location = New-Object System.Drawing.Point(20, 510)
        $DateBox.Size = New-Object System.Drawing.Size(330, 20)
        $changeForm.Controls.Add($DateBox)
    }
    if ($Track -eq "4") {
        # Manager
        $temp = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty Manager 
        $CManager = Get-ADUser $temp | Select-Object -ExpandProperty name
        $changeFormB1Text3 = New-Object System.Windows.Forms.Label
        $changeFormB1Text3.Location = '40,60'
        $changeFormB1Text3.size = '65,20'
        $changeFormB1Text3.Text = 'Manager:' 
        $changeFormB1Text4 = New-Object System.Windows.Forms.Label
        $changeFormB1Text4.Location = '140,60'
        $changeFormB1Text4.size = '200,20'
        $changeFormB1Text4.ForeColor = 'Blue'
        $changeFormB1Text4.Text = $CManager
        $changeFormBox1.Controls.AddRange(@($changeFormB1Text3, $changeFormB1Text4))
        $changeFormLable = New-Object System.Windows.Forms.Label
        $changeFormLable.Location = '50,230'
        $changeFormLable.size = '100,30'
        $changeFormLable.Text = "Manager:" 
        $changeFormcomboMan = New-Object System.Windows.Forms.ComboBox
        $changeFormcomboMan.Location = '325,230'
        $changeFormcomboMan.Size = '400,40'
        $changeFormcomboMan.AutoCompleteMode = 'Suggest'
        $changeFormcomboMan.AutoCompleteSource = 'ListItems'
        $changeFormcomboMan.Sorted = $false;
        $changeFormcomboMan.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $changeFormcomboMan.SelectedItem = $changeFormcomboMan.Items[0]
        $changeFormcomboMan.add_SelectedIndexChanged( { $changecomboBoxManager.Text = "$($changeFormcomboMan.SelectedItem.ToString())" })
        $changecomboBoxManager = New-Object System.Windows.Forms.Label
        $changecomboBoxManager.Location = '20,900'
        $changecomboBoxManager.size = '350,40'
        $changeFormcomboMan.DataSource = $Managers
        $changeForm.Controls.AddRange(@($changeFormLable, $changeFormcomboMan, $changecomboBoxManager))
    }
    elseif ($Track -eq "3") {
        $TCGrade = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty extensionAttribute5 
        if ($TCGrade -eq "1") {
            $CGrade = "AA/SGB2"
        }
        elseif ($TCGrade -eq "2") {
            $CGrade = "AO/SGB1"
        }
        elseif ($TCGrade -eq "3") {
            $CGrade = "PS"
        }
        elseif ($TCGrade -eq "4") {
            $CGrade = "EO"
        }
        elseif ($TCGrade -eq "5") {
            $CGrade = "HEO"
        }
        elseif ($TCGrade -eq "6") {
            $CGrade = "SEO"
        }
        elseif ($TCGrade -eq "7") {
            $CGrade = "GD7"
        }
        elseif ($TCGrade -eq "8") {
            $CGrade = "GD6"
        }
        else {
            $CGrade = "None"
        }
        $changeFormB1Text3 = New-Object System.Windows.Forms.Label
        $changeFormB1Text3.Location = '40,60'
        $changeFormB1Text3.size = '65,20'
        $changeFormB1Text3.Text = 'Grade:' 
        $changeFormB1Text4 = New-Object System.Windows.Forms.Label
        $changeFormB1Text4.Location = '140,60'
        $changeFormB1Text4.size = '200,20'
        $changeFormB1Text4.ForeColor = 'Blue'
        $changeFormB1Text4.Text = $CGrade
        $changeFormBox1.Controls.AddRange(@($changeFormB1Text3, $changeFormB1Text4))
        $changeFormLable = New-Object System.Windows.Forms.Label
        $changeFormLable.Location = '50,230'
        $changeFormLable.size = '100,30'
        $changeFormLable.Text = "Grade:" 
        $changeFormcomboGrade = New-Object System.Windows.Forms.ComboBox
        $changeFormcomboGrade.Location = '325,230'
        $changeFormcomboGrade.Size = '400,40'
        $changeFormcomboGrade.AutoCompleteMode = 'Suggest'
        $changeFormcomboGrade.AutoCompleteSource = 'ListItems'
        $changeFormcomboGrade.Sorted = $false;
        $changeFormcomboGrade.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $changeFormcomboGrade.SelectedItem = $changeFormcomboGrade.Items[0]
        $changeFormcomboGrade.add_SelectedIndexChanged( { $changecomboBoxGrade.Text = "$($changeFormcomboGrade.SelectedItem.ToString())" })
        $changecomboBoxGrade = New-Object System.Windows.Forms.Label
        $changecomboBoxGrade.Location = '20,900'
        $changecomboBoxGrade.size = '350,40'
        $changeFormcomboGrade.DataSource = $Lists.Grade | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
        $changeForm.Controls.AddRange(@($changeFormLable, $changeFormcomboGrade, $changecomboBoxGrade))
    }
    elseif ($Track -eq "2") {
        $COffice = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty Office
        $CTitle = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty GivenName, extensionAttribute2
        $changeFormB1Text3 = New-Object System.Windows.Forms.Label
        $changeFormB1Text3.Location = '40,60'
        $changeFormB1Text3.size = '65,20'
        $changeFormB1Text3.Text = 'Location:' 
        $changeFormB1Text4 = New-Object System.Windows.Forms.Label
        $changeFormB1Text4.Location = '140,60'
        $changeFormB1Text4.size = '200,20'
        $changeFormB1Text4.ForeColor = 'Blue'
        $changeFormB1Text4.Text = $COffice
        $changeFormBox1.Controls.AddRange(@($changeFormB1Text3, $changeFormB1Text4))
        $SelectionFormText1 = New-Object System.Windows.Forms.Label
        $SelectionFormText1.Location = '20,140'
        $SelectionFormText1.size = '350,30'
        $SelectionFormText1.Text = "Office field in AD:" 
        $SelectionFormText2 = New-Object System.Windows.Forms.Label
        $SelectionFormText2.Location = '20,175'
        $SelectionFormText2.size = '350,30'
        $SelectionFormText2.Text = "Description field in AD:" 
        $SelectionFormText3 = New-Object System.Windows.Forms.Label
        $SelectionFormText3.Location = '20,205'
        $SelectionFormText3.size = '350,30'
        $SelectionFormText3.Text = "Distribution List field in AD:" 
        $SelectionFormText4 = New-Object System.Windows.Forms.Label
        $SelectionFormText4.Location = '20,240'
        $SelectionFormText4.size = '370,30'
        $SelectionFormText4.Text = "Security Group field in AD (to access N drives):" 
        $SelectionFormcomboBox1 = New-Object System.Windows.Forms.ComboBox
        $SelectionFormcomboBox1.Location = '425,140'
        $SelectionFormcomboBox1.Size = '250,40'
        $SelectionFormcomboBox1.AutoCompleteMode = 'Suggest'
        $SelectionFormcomboBox1.AutoCompleteSource = 'ListItems'
        $SelectionFormcomboBox1.Sorted = $false;
        $SelectionFormcomboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SelectionFormcomboBox1.SelectedItem = $SelectionFormcomboBox1.Items[0]
        $SelectionFormcomboBox1.add_SelectedIndexChanged( { $SelectionFormcomboBoxOfficeSelect.Text = "$($SelectionFormcomboBox1.SelectedItem.ToString())" })
        $SelectionFormcomboBox2 = New-Object System.Windows.Forms.ComboBox
        $SelectionFormcomboBox2.Location = '425,175'
        $SelectionFormcomboBox2.Size = '250,40'
        $SelectionFormcomboBox2.AutoCompleteMode = 'Suggest'
        $SelectionFormcomboBox2.AutoCompleteSource = 'ListItems'
        $SelectionFormcomboBox2.Sorted = $false;
        $SelectionFormcomboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SelectionFormcomboBox2.SelectedItem = $SelectionFormcomboBox2.Items[0]
        $SelectionFormcomboBox2.add_SelectedIndexChanged( { $SelectionFormcomboBoxDescriptionSelect.Text = "$($SelectionFormcomboBox2.SelectedItem.ToString())" })
        $SelectionFormcomboBox3 = New-Object System.Windows.Forms.ComboBox
        $SelectionFormcomboBox3.Location = '425,205'
        $SelectionFormcomboBox3.Size = '250,40'
        $SelectionFormcomboBox3.AutoCompleteMode = 'Suggest'
        $SelectionFormcomboBox3.AutoCompleteSource = 'ListItems'
        $SelectionFormcomboBox3.Sorted = $false;
        $SelectionFormcomboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SelectionFormcomboBox3.SelectedItem = $SelectionFormcomboBox3.Items[0]
        $SelectionFormcomboBox3.add_SelectedIndexChanged( { $SelectionFormcomboBoxDistributionSelect.Text = "$($SelectionFormcomboBox3.SelectedItem.ToString())" })
        $SelectionFormcomboBox4 = New-Object System.Windows.Forms.ComboBox
        $SelectionFormcomboBox4.Location = '425,240'
        $SelectionFormcomboBox4.Size = '250,40'
        $SelectionFormcomboBox4.AutoCompleteMode = 'Suggest'
        $SelectionFormcomboBox4.AutoCompleteSource = 'ListItems'
        $SelectionFormcomboBox4.Sorted = $false;
        $SelectionFormcomboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SelectionFormcomboBox4.SelectedItem = $SelectionFormcomboBox4.Items[0]
        $SelectionFormcomboBox4.add_SelectedIndexChanged( { $SelectionFormcomboBoxSecuritySelect.Text = "$($SelectionFormcomboBox4.SelectedItem.ToString())" })
        $SelectionFormcomboBoxOfficeSelect = New-Object System.Windows.Forms.Label
        $SelectionFormcomboBoxOfficeSelect.Location = '20,600'
        $SelectionFormcomboBoxOfficeSelect.size = '350,40'
        $SelectionFormcomboBoxDescriptionSelect = New-Object System.Windows.Forms.Label
        $SelectionFormcomboBoxDescriptionSelect.Location = '20,650'
        $SelectionFormcomboBoxDescriptionSelect.size = '350,40'
        $SelectionFormcomboBoxDistributionSelect = New-Object System.Windows.Forms.Label
        $SelectionFormcomboBoxDistributionSelect.Location = '20,700'
        $SelectionFormcomboBoxDistributionSelect.size = '350,40'
        $SelectionFormcomboBoxSecuritySelect = New-Object System.Windows.Forms.Label
        $SelectionFormcomboBoxSecuritySelect.Location = '20,750'
        $SelectionFormcomboBoxSecuritySelect.size = '350,40'
        $SelectionFormText6 = New-Object System.Windows.Forms.Label
        $SelectionFormText6.Location = '20,280'
        $SelectionFormText6.size = '370,30'
        $SelectionFormText6.Text = "Department:" 
        $SelectionFormcomboBox6 = New-Object System.Windows.Forms.ComboBox
        $SelectionFormcomboBox6.Location = '425,285'
        $SelectionFormcomboBox6.Size = '250,40'
        $SelectionFormcomboBox6.AutoCompleteMode = 'Suggest'
        $SelectionFormcomboBox6.AutoCompleteSource = 'ListItems'
        $SelectionFormcomboBox6.Sorted = $false;
        $SelectionFormcomboBox6.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $SelectionFormcomboBox6.SelectedItem = $SelectionFormcomboBox6.Items[0]
        $SelectionFormcomboBox6.add_SelectedIndexChanged( { $SelectionFormcomboBoxDepartment.Text = "$($SelectionFormcomboBox6.SelectedItem.ToString())" })
        $SelectionFormcomboBoxDepartment = New-Object System.Windows.Forms.Label
        $SelectionFormcomboBoxDepartment.Location = '20,850'
        $SelectionFormcomboBoxDepartment.size = '350,40'
        $SelectionFormcomboBox6.DataSource = $Lists.Department | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
        if (($CTitle.Office -like "Sherif*") -or ($CTitle.Office -like "Lord") -or ($CTitle.Office -like "Lord") ) {
            $SelectionFormcomboBox4.DataSource = $Lists.SecurityGroup | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            $SelectionFormcomboBox3.DataSource = $Lists.distributionJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            $SelectionFormcomboBox2.DataSource = $Lists.descriptionJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            $SelectionFormcomboBox1.DataSource = $Lists.officeJud | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
        }
        else {
            if ($Title -eq "USR-PERT") {
                $SelectionFormcomboBox4.DataSource = $Lists.SecurityGroupsTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormcomboBox3.DataSource = $Lists.DistributionTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormcomboBox2.DataSource = $Lists.DescriptionTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormcomboBox1.DataSource = $Lists.OfficeTribs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
            }
            else {
                $SelectionFormcomboBox4.DataSource = $Lists.SecurityGroup | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormcomboBox3.DataSource = $Distributionlists 
                $SelectionFormcomboBox2.DataSource = $Lists.Description | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } 
                $SelectionFormcomboBox1.DataSource = $Lists.OfficeList | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            }
        }
        $changeForm.Controls.AddRange(@($SelectionFormText6, $SelectionFormText4, $SelectionFormText3, $SelectionFormText2, $SelectionFormText1))
        $changeForm.Controls.AddRange(@($SelectionFormcomboBox1, $SelectionFormcomboBox2, $SelectionFormcomboBox3, $SelectionFormcomboBox4, $SelectionFormcomboBox6, $SelectionFormcomboBoxDepartment, $SelectionFormcomboBoxGrade, $SelectionFormcomboBoxSecuritySelect, $SelectionFormcomboBoxDistributionSelect, $SelectionFormcomboBoxDescriptionSelect, $SelectionFormcomboBoxOfficeSelect))
    
    }
    else {
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(20, 165)
        $label2.Size = New-Object System.Drawing.Size(280, 20)
        $label2.Text = 'Last name:'
        $changeForm.Controls.Add($label2)
        $textBox2 = New-Object System.Windows.Forms.TextBox
        $textBox2.Location = New-Object System.Drawing.Point(20, 190)
        $textBox2.Size = New-Object System.Drawing.Size(330, 20)
        $changeForm.Controls.Add($textBox2)
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
            $changeForm.Close()
            $changeForm.Dispose()
            Return MainForm })
    $changeForm.Controls.AddRange(@($changeFormBox1, $OKButton, $CancelButton))
    $changeForm.AcceptButton = $OKButton
    $changeForm.CancelButton = $CancelButton
    $changeForm.Add_Shown( { $changeForm.Activate() })    
    $dialogResult = $changeForm.ShowDialog()
    if ($dialogResult -eq 'OK') {
        if ($Track -eq "1") {
            # rename
            $lastname = $textBox2.Text
            $SAM = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty SamAccountName, mail, GivenName, extensionattribute2, DistinguishedName 
            $firstname = $SAM.GivenName
            $Old = $SAM.mail
            $OldEmail = "smtp:$Old"
            $TempDisplayName = $LastName + ", " + $firstname
            $tentativeSAM = ($firstname.substring(0, 1) + $lastname).toLower()
            $NDisplayName = $TempDisplayName 
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
            $NSAM = $tentativeSAM
            if (Get-ADUser -Filter { displayName -eq $TempDisplayName }) {    
                do {
                    $incB ++
                    $TempDisplayName = $NDisplayName + [string]$incB
                } 
                until (-not (Get-ADUser -Filter { displayName -eq $TempDisplayName }))
            }
            $NDisplayName = $TempDisplayName
            $Proxies = @()
            $TribProxy = "smtp:$NSAM@scotcourtstribunals.gov.uk"
            $SCSPrimary = "SMTP:$NSAM@scotcourts.gov.uk"
            $TribPrimary = "SMTP:$NSAM@scotcourtstribunals.gov.uk"
            $SCSProxy = "smtp:$NSAM@scotcourts.gov.uk"
            $newProxy1 = "smtp:$NSAM@scotcourts.pnn.gov.uk"
            $newProxy2 = "smtp:$NSAM@scotcourtstribunals.pnn.gov.uk"
            $newProxy3 = "smtp:$NSAM@scotcourtsgovuk.mail.onmicrosoft.com"
            $newProxy4 = "X400:C=GB;A=CWMAIL;P=SCS;O=SCOTTISH COURTS;S=" + $Lastname + ";G=" + $FirstName + ";"
            if ($SAM.extensionattribute5 -eq "USR-PERS") {
                $Proxies = @($TribProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4, $OldEmail)                
                Set-ADUser -identity $SAM.SamAccountName -replace @{proxyAddresses = ($SCSPrimary) }
            }
            else {
                $Proxies = @($SCSProxy, $newProxy1, $newProxy2, $newProxy3, $newProxy4, $OldEmail)
                Set-ADUser -identity $SAM.SamAccountName -replace @{proxyAddresses = ($TribPrimary) }
            }
            foreach ($Proxy in $Proxies) {
                Set-ADUser -identity $SAM.SamAccountName -add @{proxyAddresses = ($Proxy) }
            }
            Rename-ADObject -Identity $SAM.DistinguishedName -NewName $newUserObjectName

            
        }
        elseif ($Track -eq "2") {
            # move
            
        }
        elseif ($Track -eq "3") {
            # Rank
            $Grade = $changecomboBoxGrade.Text
            if ($Grade -eq 'AA/SGB2') {
                $GradeNumber = "1"
            }
            elseif ($Grade -eq 'AO/SGB1') {
                $GradeNumber = "2"
            }
            elseif ($Grade -eq 'PS') {
                $GradeNumber = "3"
            }
            elseif ($Grade -eq 'EO') {
                $GradeNumber = "4"
            }
            elseif ($Grade -eq 'HEO') {
                $GradeNumber = "5"
            }
            elseif ($Grade -eq 'SEO') {
                $GradeNumber = "6"
            }
            elseif ($Grade -eq 'GD7') {
                $GradeNumber = "7"
            }
            elseif ($Grade -eq 'GD6') {
                $GradeNumber = "8"
            }
            else {
                Write-Output "No Change"
                $changeForm.Close() | Out-Null
                MainForm
            }
            $SAM = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty SamAccountName
            Set-ADUser -Identity $SAM -clear "extensionattribute5"
            Set-ADUser -Identity $SAM -add @{"extensionattribute5" = $GradeNumber }
            $changeForm.Close() | Out-Null
            MainForm
        }
        elseif ($Track -eq "4") {
            # Manager
            $TManager = $changecomboBoxManager.Text
            $SAM = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty SamAccountName
            $Manager = Get-ADUser -Filter "Displayname -eq '$TManager'" -Properties * | Select-Object -ExpandProperty DistinguishedName
            Set-ADUser -Identity $SAM -Manager $Manager
            $changeForm.Close() | Out-Null
            MainForm
        }
        elseif ($Track -eq "5") {
            # Contractor
            $Enddate = $DateBox.Text
            $SAM = Get-ADUser -Filter "Displayname -eq '$DisplayName'" -Properties * | Select-Object -ExpandProperty SamAccountName
            Set-ADAccountExpiration -Identity $SAM -DateTime $Enddate
            $changeForm.Close() | Out-Null
            MainForm
        }
    }
}
Function MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [int]$incA = 0
    [int]$incB = 0
    $Track = "0"
    $ManForm = New-Object System.Windows.Forms.Form
    $ManForm.Icon = $Icon
    $ManForm.Size = New-Object System.Drawing.Size(375, 475)
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
    $OKButton.Location = New-Object System.Drawing.Point(175, 375)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ManForm.AcceptButton = $OKButton
    $ManForm.Controls.Add($OKButton)
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(250, 375)
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
    $label1.Text = 'Select User:'
    $ManForm.Controls.Add($label1)
    $ComboBox1 = New-Object System.Windows.Forms.ComboBox
    $ComboBox1.Location = '20,160'
    $ComboBox1.Size = '300, 310'
    $ComboBox1.AutoCompleteMode = 'Suggest'
    $ComboBox1.AutoCompleteSource = 'ListItems'
    $ComboBox1.Sorted = $true;
    $ComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $ComboBox1.DataSource = $UsernameList
    $ManForm.Controls.Add($ComboBox1)
    $RenameButton = New-Object System.Windows.Forms.RadioButton
    $RenameButton.Location = '20,200'
    $RenameButton.size = '200,30'
    $RenameButton.Checked = $true 
    $RenameButton.Text = 'Rename.'
    $MoveButton = New-Object System.Windows.Forms.RadioButton
    $MoveButton.Location = '20,230'
    $MoveButton.size = '200,30'
    $MoveButton.Checked = $false
    $MoveButton.Text = 'Move.'
    $RankButton = New-Object System.Windows.Forms.RadioButton
    $RankButton.Location = '20,260'
    $RankButton.size = '200,30'
    $RankButton.Checked = $false
    $RankButton.Text = 'Grade.'
    $ManaButton = New-Object System.Windows.Forms.RadioButton
    $ManaButton.Location = '20,290'
    $ManaButton.size = '200,30'
    $ManaButton.Checked = $false
    $ManaButton.Text = 'Manager change.'
    $ContButton = New-Object System.Windows.Forms.RadioButton
    $ContButton.Location = '20,320'
    $ContButton.size = '200,30'
    $ContButton.Checked = $false
    $ContButton.Text = 'Contractor change.'
    $ManForm.Controls.Add($RenameButton)
    $ManForm.Controls.Add($MoveButton)
    $ManForm.Controls.Add($ManaButton)
    $ManForm.Controls.Add($RankButton)
    $ManForm.Controls.Add($ContButton)
    $ManForm.Topmost = $true
    $result = $ManForm.ShowDialog()
    if ($Result -eq 'OK') {
        $DisplayName = $ComboBox1.Text
        $SAM = Get-ADUser -Filter "Displayname -eq '$DisplayName'" | Select-Object -ExpandProperty 'SamAccountName'
        if ($RenameButton.Checked) {
            $Track = "1"
            change
        }
        elseif ($MoveButton.Checked) {
            $Track = "2"
            change
        }
        elseif ($RankButton.Checked) {
            $Track = "3"
            change
        }
        elseif ($ManaButton.Checked) {
            $Track = "4"
            change
        }
        elseif ($ContButton.Checked) {
            $Track = "5"
            change
        }
    }
}

