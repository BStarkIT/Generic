<#
.SYNOPSIS
This PowerShell script is to create shared mailboxes.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>

$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\RBAC\$Date.txt" -append
$version = "1.0"
$UserName = $env:username
$DC = "SAU-DC-04.scotcourts.local"
$Attrib2 = "ADM-PERS"
$OU = "OU=RBAC Admins,OU=User Accounts (Admin),OU=SCTS,DC=scotcourts,DC=local"
$WinTitle = "Create RBAC v$version."
if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
}
else {
    Function MainForm {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $UserNameList = Get-ADUser -filter * -searchbase 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
        Write-Output "Pulling User Accounts"
        $RBACGroups = Get-ADGroup -filter * -searchbase 'OU=RBAC,OU=Groups,OU=SCTS,DC=scotcourts,DC=local' -Properties Name | Select-Object name | Select-Object -ExpandProperty name
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
        $OKButton.Location = New-Object System.Drawing.Point(175, 400)
        $OKButton.Size = New-Object System.Drawing.Size(75, 23)
        $OKButton.Text = 'OK'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ManForm.AcceptButton = $OKButton
        $ManForm.Controls.Add($OKButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(250, 400)
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
        $label1.Text = 'Ticket:'
        $ManForm.Controls.Add($label1)
        $textBox1 = New-Object System.Windows.Forms.TextBox
        $textBox1.Location = New-Object System.Drawing.Point(10, 150)
        $textBox1.Size = New-Object System.Drawing.Size(280, 20)
        $ManForm.Controls.Add($textBox1)
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(10, 180)
        $label2.Size = New-Object System.Drawing.Size(280, 20)
        $label2.text = "Select the User to create RBAC account:"
        $ManForm.Controls.Add($label2)
        $UserCombo = New-Object System.Windows.Forms.ComboBox
        $UserCombo.Location = '10,210'
        $UserCombo.Size = '280,40'
        $UserCombo.AutoCompleteMode = 'Suggest'
        $UserCombo.AutoCompleteSource = 'ListItems'
        $UserCombo.Sorted = $True;
        $UserCombo.Enabled = $True;
        $UserCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $UserCombo.SelectedItem = $UserCombo.Items[0]
        $UserCombo.DataSource = $UserNameList
        $UserCombo.add_SelectedIndexChanged( { $NameSelect.Text = "$($UserCombo.SelectedItem.ToString())" })
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10, 240)
        $label3.Size = New-Object System.Drawing.Size(280, 20)
        $label3.text = "Select the RBAC Group:"
        $ManForm.Controls.Add($label3)
        $GroupCombo = New-Object System.Windows.Forms.ComboBox
        $GroupCombo.Location = '10,275'
        $GroupCombo.Size = '280,40'
        $GroupCombo.AutoCompleteMode = 'Suggest'
        $GroupCombo.AutoCompleteSource = 'ListItems'
        $GroupCombo.Sorted = $True;
        $GroupCombo.Enabled = $True;
        $GroupCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $GroupCombo.SelectedItem = $UserCombo.Items[0]
        $GroupCombo.DataSource = $RBACGroups
        $GroupCombo.add_SelectedIndexChanged( { $groupSelect.Text = "$($GroupCombo.SelectedItem.ToString())" })
        $NameSelect = New-Object System.Windows.Forms.Label
        $NameSelect.Location = '40,330'
        $NameSelect.size = '200,20'
        $Selected = New-Object System.Windows.Forms.Label
        $Selected.Location = '30,310'
        $Selected.size = '100,30'
        $Selected.Text = "Selected:" 
        $groupSelect = New-Object System.Windows.Forms.Label
        $groupSelect.Location = '40,350'
        $groupSelect.size = '350,40'
        $ManForm.Controls.AddRange(@($UserCombo, $GroupCombo, $NameSelect, $groupSelect, $Selected))
        $ManForm.Topmost = $true
        $ManForm.Add_Shown( { $textBox1.Select() })
        $result = $ManForm.ShowDialog()
        $tempTicket = $textBox1.Text
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
            $pwdList += $charlist1 | Get-Random
            $pwdList += $charlist2 | Get-Random
            $pwdList += $charlist3 | Get-Random
        }
        $pass = -join ($pwdList | get-random -count $pwdList.count)
        Write-Host "Please Note: This account will be made with Password: " $pass -ForegroundColor Red
        $password = ConvertTo-SecureString $pass -AsPlainText -Force
        if ($Result -eq 'OK') {
            $Ticket = $tempTicket -replace '[^\x30-\x39]+', ''
            if ($null -eq  $Ticket) {
                Write-Host "No Ticket number entered"
                [System.Windows.Forms.MessageBox]::Show("You need to enter a Ticket number!  Trying to enter blank fields is never a good idea.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            if ($textBox1.Text -eq '') {
                Write-Host "No Ticket number entered"
                [System.Windows.Forms.MessageBox]::Show("You need to enter a Ticket number!  Trying to enter blank fields is never a good idea.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            elseif (($NameSelect.Text -eq '') -or ($groupSelect.Text -eq '')) {
                Write-Host "No details selected"
                [System.Windows.Forms.MessageBox]::Show("You need to select a user & group.", $WinTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $ManForm.Close()
                $ManForm.Dispose()
                break
            }
            else {
                $Group = $groupSelect.text
                $Name = $NameSelect.text + " - RBAC"
                $Base = (Get-ADUser -filter { DisplayName -eq $NameSelect.Text } -Properties * | Select-Object SamAccountName, GivenName, Surname)
                $FirstName = $Base.GivenName
                $Surname = $Base.Surname
                $Sam1 = $Base.SamAccountName
                $Sam = $Sam1 + "_a"
                $UPN = $Sam + "@scotcourts.gov.uk"
                $Description = $Group + " - " + $Ticket
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
                New-AdUser -Name $Name -SamAccountName $Sam -GivenName $FirstName -Surname $Surname -DisplayName $Name -UserPrincipalName $UPN -Description $Description -Path $OU -Enabled $True -ChangePasswordAtLogon $false -Server $DC -AccountPassword $Password -passThru
                Start-Sleep -Seconds 5
                Set-ADUser -Identity $SAM -add @{"extensionattribute2" = $Attrib2 }
                Set-ADUser -Identity $SAM -add @{"extensionattribute3" = $Ticket }
                Add-ADGroupMember -Identity $Group -Members $SAM
                $copy = "On ticket $Ticket user $Sam made & added to group $Group Password set as $pass " | clip
                Write-Host "On ticket $Ticket user $Sam made & added to group $Group Password set as $pass "
                Pause

            }
        }
    }
    Return MainForm
}
