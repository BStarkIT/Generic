$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\email.ico'
$Version = '1.00'
$WinTitle = "Office 365 control script v$version."
$Date = Get-Date -Format "dd-MM-yyyy"
Start-Transcript -Path "\\scotcourts.local\data\CDiScripts\Scripts\Logs\Exchange\$Date.txt" -append
Connect-ExchangeOnline
Write-Output "Pulling Shared mailboxes"
$Sharedmailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -Properties DisplayName | Select-Object DisplayName | Select-Object -ExpandProperty DisplayName 
Write-Output "Pulling Distribution groups"
$Distributionlists = Get-DistributionGroup | Select-Object Name | Select-Object -ExpandProperty Name  
Write-Output "Pulling User Accounts"
$UserNameList = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$SMBOU = 'OU=Shared Mailboxes,OU=Resource Accounts,OU=SCTS,DC=scotcourts,DC=local'
$DistOU = 'OU=Distribution Lists,OU=Groups,OU=SCTS,DC=scotcourts,DC=local'
$DC = "SAU-DC-04.scotcourts.local"

Function Detailsform {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $Detailsform = New-Object System.Windows.Forms.Form
    $Detailsform.width = 750
    $Detailsform.height = 450
    $Detailsform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Detailsform.MinimizeBox = $False
    $Detailsform.MaximizeBox = $False
    $Detailsform.FormBorderStyle = 'Fixed3D'
    $Detailsform.Icon = $Icon
    $Detailsform.Font = New-Object System.Drawing.Font('Ariel', 10)
    $DetailsformBox1 = New-Object System.Windows.Forms.GroupBox
    $DetailsformBox1.Location = '40,30'
    $DetailsformBox1.size = '650,300'
    $Mailbox = New-Object System.Windows.Forms.Label
    $Mailbox.Location = '400,80'
    $Username = New-Object System.Windows.Forms.Label
    $Username.Location = '400,80'
    $Dist = New-Object System.Windows.Forms.Label
    $Dist.Location = '400,80'
    # 11 SMB - Check access
    if ($selection -eq "11") {
        $Detailsform.Text = "check Shared mailbox access"
        $DetailsformBox1.text = 'Select a MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,80'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'MailBoxName:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,75'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $Sharedmailboxes 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1))
    }
    # 12 SMB - Add access
    elseif ($selection -eq "12") {
        $Detailsform.Text = "Add Shared mailbox access"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    # 13 SMB - Send
    elseif ($selection -eq "13") {
        $Detailsform.Text = "Add Shared mailbox access with send on behalf"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    # 14 SMB - Remove access SMB
    elseif ($selection -eq "14") {
        $Detailsform.Text = "Remove Shared mailbox access"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    # 21 OOO check
    elseif ($selection -eq "21") {
        $Detailsform.Text = "Check OOO"
        $DetailsformBox1.text = 'Select a User from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $Username))
    }
    # 22 OOO add
    elseif ($selection -eq "22") {
        $Detailsform.Text = "Add OOO"
        $DetailsformBox1.text = 'Select a User from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,20'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,70'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'OOO Message:'
        $textBox1 = New-Object System.Windows.Forms.TextBox
        $textBox1.Location = '50,110'
        $textBox1.Size = '450,75'
        $textBox1.Multiline = $true
        $textBox1.AcceptsReturn = $true
        $textBox1.WordWrap = $true
        $textBox1.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $textBox1, $Username))
    }
    # 23 OOO off
    elseif ($selection -eq "23") {
        $Detailsform.Text = "Turn off OOO"
        $DetailsformBox1.text = 'Select a User from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $Username))
    }
    # 31 Cal Check
    elseif ($selection -eq "31") {
        $Detailsform.Text = "check Calendar access"
        $DetailsformBox1.text = 'Select a MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,80'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'MailBoxName:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,75'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $Sharedmailboxes 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $Mailbox))
    }
    # 32 cal Reviewer
    elseif ($selection -eq "32") {
        $Detailsform.Text = "Add Calendar Reviewer access"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    # 33 cal Owner
    elseif ($selection -eq "33") {
        $Detailsform.Text = "Add Calendar Owner access"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    # 34 Cal remove
    elseif ($selection -eq "34") {
        $Detailsform.Text = "Remove Calendar access"
        $DetailsformBox1.text = 'Select a User and MailBoxName from the dropdown lists:.'
        $DetailsformMenuLable1 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable1.Location = '20,40'
        $DetailsformMenuLable1.size = '150,40'
        $DetailsformMenuLable1.Text = 'User:'
        $DetailsformMenu1 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu1.Location = '225,35'
        $DetailsformMenu1.Size = '350, 350'
        $DetailsformMenu1.AutoCompleteMode = 'Suggest'
        $DetailsformMenu1.AutoCompleteSource = 'ListItems'
        $DetailsformMenu1.Sorted = $true;
        $DetailsformMenu1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu1.DataSource = $UserNameList 
        $DetailsformMenu1.add_SelectedIndexChanged( { $Username.Text = "$($DetailsformMenu1.SelectedItem.ToString())" })
        $DetailsformMenuLable2 = New-Object System.Windows.Forms.Label
        $DetailsformMenuLable2.Location = '20,80'
        $DetailsformMenuLable2.size = '150,40'
        $DetailsformMenuLable2.Text = 'MailBoxName:'
        $DetailsformMenu2 = New-Object System.Windows.Forms.ComboBox
        $DetailsformMenu2.Location = '225,75'
        $DetailsformMenu2.Size = '350, 350'
        $DetailsformMenu2.AutoCompleteMode = 'Suggest'
        $DetailsformMenu2.AutoCompleteSource = 'ListItems'
        $DetailsformMenu2.Sorted = $true;
        $DetailsformMenu2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $DetailsformMenu2.DataSource = $Sharedmailboxes 
        $DetailsformMenu2.add_SelectedIndexChanged( { $Mailbox.Text = "$($DetailsformMenu2.SelectedItem.ToString())" })
        $DetailsformBox1.Controls.AddRange(@($DetailsformMenuLable1, $DetailsformMenu1, $DetailsformMenuLable2, $DetailsformMenu2, $Mailbox, $Username))
    }
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,350'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,350'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm })
    $Detailsform.Controls.AddRange(@($DetailsformBox1, $OKButton, $CancelButton))
    $Detailsform.AcceptButton = $OKButton
    $Detailsform.CancelButton = $CancelButton
    $Detailsform.Add_Shown( { $Detailsform.Activate() })    
    $Result = $Detailsform.ShowDialog()
    if ($Result -eq 'OK') {
        if ($Selection -eq "11") {
            # 11 SMB - Check access
            $SMB = $Mailbox.Text 
            $Mailboxname = Get-EXOMailbox -Identity $SMB | Select-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress 
            $status = Get-EXOMailboxPermission -Identity $Mailboxname | Where-Object { $_.AccessRights -eq 'FullAccess' } | Where-Object { $_.user -notlike "s-1*" -and @("scs\domain admins", "scs\enterprise admins", "nt authority\system", "scs\organization management") -notcontains $_.User } | Select-Object User, AccessRights
            $status | Select-Object @{label = 'User'; expression = { $_.User -replace '^SCS\\' } } | Sort-Object user | Out-GridView -Title "List of Users with Full Access permissions on $MailboxName mailbox" -Wait 
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "12") {
            # 12 SMB - Add access
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text
            $PUser = $Username.Text
            Add-MailboxPermission -Identity $Mailboxname.Name -User $UsernameSAM.PrimarySmtpAddress -AccessRights FullAccess -InheritanceType All
            Write-output "User $PUser add to $PMailbox mailbox"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "13") {
            # 13 SMB - Send
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text
            $PUser = $Username.Text
            Add-MailboxPermission -Identity $Mailboxname.Name -User $UsernameSAM.PrimarySmtpAddress -AccessRights FullAccess -InheritanceType All
            Set-Mailbox -Identity $Mailboxname.Name -GrantSendOnBehalfTo @{add = $UsernameSAM.PrimarySmtpAddress }
            Write-output "User $PUser add to $PMailbox mailbox with Send on behalf permission"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "14") {
            # 14 SMB - Remove from
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text
            $PUser = $Username.Text
            Remove-MailboxPermission -Identity $Mailboxname.Name -user $UsernameSAM.PrimarySmtpAddress -AccessRights FullAccess -InheritanceType All -confirm:$false
            Set-Mailbox -Identity $Mailboxname.Name -GrantSendOnBehalfTo @{remove = $UsernameSAM.PrimarySmtpAddress }
            Write-output "User $PUser removed from $PMailbox mailbox"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "21") {
            # 21 OOO check
            $PUser = $Username.Text
            $PrimarySmtpAddress = get-exomailbox $Username.Text | Select-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress
            $Status = Get-MailboxAutoReplyConfiguration $PrimarySmtpAddress | Select-Object AutoReplyState  
            If ($Status.AutoReplyState -eq 'Disabled') {
                $PStatus = "Disabled"
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The Out of Office for user $($DetailsformMenu1.SelectedItem.ToString()) is currently:                      DISABLED - turned off.", 'User Mailbox - Out Of Office - Check current status.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            ElseIf ($Status.AutoReplyState -eq 'Enabled') {
                $PStatus = "Enabled"
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The Out of Office for use $($DetailsformMenu1.SelectedItem.ToString()) is currently:                       ENABLED - turned on.", 'User Mailbox - Out Of Office - Check current status.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            Write-Output "The $PUser Out of Office is currently $PStatus"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "22") {
            # 22 OOO add
            $PUser = $Username.Text
            Set-MailboxAutoReplyConfiguration $PrimarySmtpAddress -AutoreplyState Enabled -InternalMessage "$($textBox1.Text)" -ExternalMessage "$($textBox1.Text)"
            Write-Output "The $PUser Out of Office is Now Enabled"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "23") {
            # 23 OOO off
            $PUser = $Username.Text
            $PrimarySmtpAddress = get-exomailbox $Username.Text | Select-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress
            Set-MailboxAutoReplyConfiguration $PrimarySmtpAddress -AutoreplyState Disabled
            Write-Output "The $PUser Out of Office is Now Disabled"
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "31") {
            # 31 Cal Check
            Write-Output $Mailbox.Text
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text | Select-Object -ExpandProperty UserPrincipalName
            Write-Output $Mailboxname
            Get-EXOMailboxFolderPermission -Identity ($Mailboxname + ':\calendar') |  Select-Object User, AccessRights | Sort-Object user | Out-GridView -Title "List of Users with permissions on $Mailboxname calendar" -Wait 
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "32") {
            # 32 cal Reviewer
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text | Select-Object -ExpandProperty UserPrincipalName
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text | Select-Object -ExpandProperty PrimarySmtpAddress
            $PUser = $Username.Text
            $Status = Get-EXOMailboxFolderPermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM
            If ($Status.AccessRights -eq 'Reviewer') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user ( $($DetailsformMenu1.SelectedItem.ToString()) ) already has Reviewer Permissions to the ( $($DetailsformMenu2.SelectedItem.ToString()) ) calendar.", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseIf ($Status.AccessRights -eq 'Owner') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user ( $($DetailsformMenu1.SelectedItem.ToString()) ) already has OWNER Permissions to the ( $($DetailsformMenu1.SelectedItem.ToString()) ) calendar. Remove the OWNER permissions before adding Reviewer permissions", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            else {
                Add-MailboxFolderPermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM -AccessRights Reviewer
                Write-output "User $PUser Added to $PMailbox Calendar as Reviewer"
            }
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "33") {
            # 33 cal Owner
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text | Select-Object -ExpandProperty UserPrincipalName
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text | Select-Object -ExpandProperty PrimarySmtpAddress
            $PUser = $Username.Text
            $Status = Get-EXOMailboxFolderPermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM
            If ($Status.AccessRights -eq 'Reviewer') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user $PUser already has Reviewer Permissions to the $PMailbox calendar.", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            elseIf ($Status.AccessRights -eq 'Owner') {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user $PUser already has OWNER Permissions to the $PMailbox calendar. Remove the OWNER permissions before adding Reviewer permissions", "Calendar - Add Reviewer Permissions for a User", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            else {
                Add-MailboxFolderPermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM -AccessRights Owner
                Write-output "User $PUser Added to $PMailbox Calendar as Owner"
            }
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
        elseif ($Selection -eq "34") {
            # 34 Cal remove
            $Mailboxname = Get-EXOMailbox -Identity $Mailbox.Text | Select-Object -ExpandProperty UserPrincipalName
            $PMailbox = $Mailbox.Text
            $UsernameSAM = Get-EXOMailbox -Identity $Username.Text | Select-Object -ExpandProperty PrimarySmtpAddress
            $PUser = $Username.Text
            try {
                $Status = Get-EXOMailboxFolderPermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM
                Remove-mailboxfolderpermission -identity ($Mailboxname + ':\calendar') -user $UsernameSAM -confirm:$false
                Write-output "User $PUser Removed from $PMailbox Calendar"
            }
            catch {
                Add-Type -AssemblyName System.Windows.Forms 
                [System.Windows.Forms.MessageBox]::Show("The user $PUser does not have Calendar Access Permissions to the $PMailbox Calendar.", 'Calendar - Remove Calendar Access Permissions for a User', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
            $Detailsform.Close()
            $Detailsform.Dispose()
            MainForm
        }
    
    }
}

Function Selectionform {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $Selectionform = New-Object System.Windows.Forms.Form
    $Selectionform.width = 750
    $Selectionform.height = 450
    $Selectionform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Selectionform.MinimizeBox = $False
    $Selectionform.MaximizeBox = $False
    $Selectionform.FormBorderStyle = 'Fixed3D'
    $Selectionform.Text = $WinTitle
    $Selectionform.Icon = $Icon
    $Selectionform.Font = New-Object System.Drawing.Font('Ariel', 10)
    $SelectionformBox1 = New-Object System.Windows.Forms.GroupBox
    $SelectionformBox1.Location = '40,30'
    $SelectionformBox1.size = '650,300'
    $SelectionformBox1.text = 'Select an option.'
    if ($Selection -eq "1") {
        $SelectionformRadio1 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio1.Location = '20,40'
        $SelectionformRadio1.size = '600,40'
        $SelectionformRadio1.Checked = $true 
        $SelectionformRadio1.Text = 'Shared Mailbox - Check access:'
        $SelectionformRadio2 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio2.Location = '20,80'
        $SelectionformRadio2.size = '600,40'
        $SelectionformRadio2.Checked = $false 
        $SelectionformRadio2.Text = 'Shared Mailbox - Add access:'
        $SelectionformRadio3 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio3.Location = '20,120'
        $SelectionformRadio3.size = '600,40'
        $SelectionformRadio3.Checked = $false
        $SelectionformRadio3.Text = 'Shared Mailbox - Add access & Send:'
        $SelectionformRadio4 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio4.Location = '20,160'
        $SelectionformRadio4.size = '600,40'
        $SelectionformRadio4.Checked = $false
        $SelectionformRadio4.Text = 'Shared Mailbox - remove access:'
        $SelectionformBox1.Controls.AddRange(@($SelectionformRadio1, $SelectionformRadio2, $SelectionformRadio3, $SelectionformRadio4))
    }
    elseif ($Selection -eq "2") {
        $SelectionformRadio1 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio1.Location = '20,40'
        $SelectionformRadio1.size = '600,40'
        $SelectionformRadio1.Checked = $true 
        $SelectionformRadio1.Text = 'Out Of Office management - Check OOO:'
        $SelectionformRadio2 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio2.Location = '20,80'
        $SelectionformRadio2.size = '600,40'
        $SelectionformRadio2.Checked = $false 
        $SelectionformRadio2.Text = 'Out Of Office management - Add OOO:'
        $SelectionformRadio3 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio3.Location = '20,120'
        $SelectionformRadio3.size = '600,40'
        $SelectionformRadio3.Checked = $false
        $SelectionformRadio3.Text = 'Out Of Office management - Turn off OOO:'
        $SelectionformBox1.Controls.AddRange(@($SelectionformRadio1, $SelectionformRadio2, $SelectionformRadio3))
    }
    else {
        $SelectionformRadio1 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio1.Location = '20,40'
        $SelectionformRadio1.size = '600,40'
        $SelectionformRadio1.Checked = $true 
        $SelectionformRadio1.Text = 'Calendar - Check Access:'
        $SelectionformRadio2 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio2.Location = '20,80'
        $SelectionformRadio2.size = '600,40'
        $SelectionformRadio2.Checked = $false 
        $SelectionformRadio2.Text = 'Calendar - Add Reviewer Access:'
        $SelectionformRadio3 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio3.Location = '20,120'
        $SelectionformRadio3.size = '600,40'
        $SelectionformRadio3.Checked = $false
        $SelectionformRadio3.Text = 'Calendar - Add Owner Access:'
        $SelectionformRadio4 = New-Object System.Windows.Forms.RadioButton
        $SelectionformRadio4.Location = '20,160'
        $SelectionformRadio4.size = '600,40'
        $SelectionformRadio4.Checked = $false
        $SelectionformRadio4.Text = 'Calendar - remove access:'
        $SelectionformBox1.Controls.AddRange(@($SelectionformRadio1, $SelectionformRadio2, $SelectionformRadio3, $SelectionformRadio4))
    }
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '500,350'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '375,350'
    $CancelButton.Size = '100,40'
    $CancelButton.Font = New-Object System.Drawing.Font('Ariel', 7)
    $CancelButton.Text = 'Cancel back to MainForm'
    $CancelButton.add_Click( {
            $Selectionform.Close()
            $Selectionform.Dispose()
            MainForm })
    $Selectionform.Controls.AddRange(@($SelectionformBox1, $OKButton, $CancelButton))
    $Selectionform.AcceptButton = $OKButton
    $Selectionform.CancelButton = $CancelButton
    $Selectionform.Add_Shown( { $Selectionform.Activate() })    
    $Result = $Selectionform.ShowDialog()
    if ($Result -eq 'OK') {
        if ($Selection -eq "1") {
            if ($SelectionformRadio1.Checked) { 
                $selection = "11" # Shared Mailbox - Check access
            }
            elseif ($SelectionformRadio2.Checked) { 
                $selection = "12" # Shared Mailbox - add access
            }
            elseif ($SelectionformRadio3.Checked) { 
                $selection = "13" # Shared Mailbox - add + Send access
            }
            elseif ($SelectionformRadio4.Checked) { 
                $selection = "14" # Shared Mailbox - remove access
            }
        }
        elseif ($Selection -eq "2") {
            if ($SelectionformRadio1.Checked) { 
                $selection = "21" # Check OOO
            }
            elseif ($SelectionformRadio2.Checked) { 
                $selection = "22" # add OOO
            }
            elseif ($SelectionformRadio3.Checked) { 
                $selection = "23" # turn off OOO
            }
        }
        else {
            if ($SelectionformRadio1.Checked) { 
                $selection = "31" # Check Calendar
            }
            elseif ($SelectionformRadio2.Checked) { 
                $selection = "32" # Add Reviewer
            }
            elseif ($SelectionformRadio3.Checked) { 
                $selection = "33" # Add owner
            }
            elseif ($SelectionformRadio4.Checked) { 
                $selection = "34" # Remove
            }
        }
        Detailsform
    }
}

Function MainForm {
    $Selection = "0"
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    $O365Mainform = New-Object System.Windows.Forms.Form
    $O365Mainform.width = 780
    $O365Mainform.height = 550
    $O365Mainform.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $O365Mainform.MinimizeBox = $False
    $O365Mainform.MaximizeBox = $False
    $O365Mainform.FormBorderStyle = 'Fixed3D'
    $O365Mainform.Text = 'Exchange 2013 Management Main Form v2.00'
    $O365Mainform.Icon = $Icon
    $Logo = [System.Drawing.Image]::Fromfile('\\scotcourts.local\data\CDiScripts\Scripts\Resources\icons\SCTS.png')
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Width = $Logo.Size.Width
    $pictureBox.Height = $Logo.Size.Height
    $pictureBox.Image = $Logo
    $O365Mainform.controls.add($pictureBox)
    $O365Mainform.Font = New-Object System.Drawing.Font('Ariel', 10)
    $O365MainformGroupBox = New-Object System.Windows.Forms.GroupBox
    $O365MainformGroupBox.Location = '40,100'
    $O365MainformGroupBox.size = '700,320'
    $O365MainformGroupBox.text = 'Select an option.'
    $O365MainformGroupBoxRadioButton1 = New-Object System.Windows.Forms.RadioButton
    $O365MainformGroupBoxRadioButton1.Location = '20,30'
    $O365MainformGroupBoxRadioButton1.size = '600,40'
    $O365MainformGroupBoxRadioButton1.Checked = $true 
    $O365MainformGroupBoxRadioButton1.Text = 'Shared Mailbox - management.'
    $O365MainformGroupBoxRadioButton2 = New-Object System.Windows.Forms.RadioButton
    $O365MainformGroupBoxRadioButton2.Location = '20,70'
    $O365MainformGroupBoxRadioButton2.size = '600,40'
    $O365MainformGroupBoxRadioButton2.Checked = $false
    $O365MainformGroupBoxRadioButton2.Text = 'User Mailbox - Out Of Office management.'
    $O365MainformGroupBoxRadioButton3 = New-Object System.Windows.Forms.RadioButton
    $O365MainformGroupBoxRadioButton3.Location = '20,110'
    $O365MainformGroupBoxRadioButton3.size = '600,40'
    $O365MainformGroupBoxRadioButton3.Checked = $false
    $O365MainformGroupBoxRadioButton3.Text = 'Shared Calendar - management.'
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '640,450'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '525,450'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $O365Mainform.Close()
            $O365Mainform.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    $O365Mainform.Controls.AddRange(@($O365MainformGroupBox, $OKButton, $CancelButton))
    $O365MainformGroupBox.Controls.AddRange(@($O365MainformGroupBoxRadioButton1, $O365MainformGroupBoxRadioButton2, $O365MainformGroupBoxRadioButton3, $O365MainformGroupBoxRadioButton4))
    $O365Mainform.AcceptButton = $OKButton
    $O365Mainform.CancelButton = $CancelButton
    $O365Mainform.Add_Shown( { $O365Mainform.Activate() })    
    $Result = $O365Mainform.ShowDialog()
    if ($Result -eq 'OK') {
        if ($O365MainformGroupBoxRadioButton1.Checked) {
            $Selection = "1"
            Selectionform
        }
        elseif ($O365MainformGroupBoxRadioButton2.Checked) {
            $Selection = "2"
            Selectionform
        }
        elseif ($O365MainformGroupBoxRadioButton3.Checked) {
            $Selection = "3"
            Selectionform
        }
    }
}
Return MainForm
