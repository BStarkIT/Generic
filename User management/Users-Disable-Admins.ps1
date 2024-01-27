# Disable Admin Accounts  V1
# Author        Brian Stark
# Date          04/06/2021
# Version       1.0
# Purpose       To disable Admin accounts
#
# Changes       1.0 04/06/2021  BS  Forked from Leavers disable script
# Script function: This script performs the following on a User account.

$message = "
AD account - Disables admin account.
AD account - Moves user account to Z-Disabled_AdminRemoves OU.
AD account - Labels account to be deleted in 1 month."
#
#
$version = '1.0'
$Title = "Disable Admin Accounts v$version."
$UserNameList = Get-aduser -Filter { name -like "*Admin*" } -Properties DisplayName | Select-Object Displayname | select-Object -ExpandProperty DisplayName
#  Show Start Message:   #
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$StartMessage = [System.Windows.Forms.MessageBox]::Show("This script performs the following on a User account." + "`n" + "$message", $Title,
    [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {
    Exit
} 
else {
        #
    # create the pop up information form
        #
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = "Disable Admin Accounts."
        $PopForm.Size = New-Object System.Drawing.Size(420, 200)
        $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $PopUpLabel = New-Object System.Windows.Forms.Label
        $PopUpLabel.Location = '80, 40' 
        $PopUpLabel.Size = '300, 80'
        $PopUpLabel.Text = $poplabel
        $PopForm.Controls.Add($PopUpLabel)
        $PopForm.Show() | Out-Null
        Start-Sleep -Seconds 2
        $PopForm.Close() | Out-Null
    }
    #
    # create the pop up information form complete
    #
    # create the main form
    #
    Function AdminRemove {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $AdminRemoveForm = New-Object System.Windows.Forms.Form
        $AdminRemoveForm.width = 745
        $AdminRemoveForm.height = 475
        $AdminRemoveForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $AdminRemoveForm.Controlbox = $false
        $AdminRemoveForm.Icon = $Icon
        $AdminRemoveForm.FormBorderStyle = 'Fixed3D'
        $AdminRemoveForm.Text = $Title
        $AdminRemoveForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $AdminRemoveBox1 = New-Object System.Windows.Forms.GroupBox
        $AdminRemoveBox1.Location = '40,20'
        $AdminRemoveBox1.size = '650,80'
        $AdminRemoveBox1.text = '1. Select a UserName from the list:'
        $AdminRemovetextLabel1 = New-Object 'System.Windows.Forms.Label';
        $AdminRemovetextLabel1.Location = '20,35'
        $AdminRemovetextLabel1.size = '200,40'
        $AdminRemovetextLabel1.Text = 'Username:' 
        $AdminRemoveMBNameComboBox1 = New-Object System.Windows.Forms.ComboBox
        $AdminRemoveMBNameComboBox1.Location = '275,30'
        $AdminRemoveMBNameComboBox1.Size = '350, 350'
        $AdminRemoveMBNameComboBox1.AutoCompleteMode = 'Suggest'
        $AdminRemoveMBNameComboBox1.AutoCompleteSource = 'ListItems'
        $AdminRemoveMBNameComboBox1.Sorted = $true;
        $AdminRemoveMBNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $AdminRemoveMBNameComboBox1.DataSource = $UserNameList
        $AdminRemoveMBNameComboBox1.Add_SelectedIndexChanged( { $AdminRemoveSelectedMailBoxNametextLabel6.Text = "$($AdminRemoveMBNameComboBox1.SelectedItem.ToString())" })
        $AdminRemoveBox2 = New-Object System.Windows.Forms.GroupBox
        $AdminRemoveBox2.Location = '40,120'
        $AdminRemoveBox2.size = '650,215'
        $AdminRemoveBox2.text = '2. Check the details below are correct before proceeding:'
        $AdminRemovetextLabel3 = New-Object System.Windows.Forms.Label
        $AdminRemovetextLabel3.Location = '20,40'
        $AdminRemovetextLabel3.size = '100,40'
        $AdminRemovetextLabel3.Text = 'The User:'
        $AdminRemovetextLabel4 = New-Object System.Windows.Forms.Label
        $AdminRemovetextLabel4.Location = '330,20'
        $AdminRemovetextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $AdminRemovetextLabel4.size = '310,170'
        $AdminRemovetextLabel4.Text = $message
        $AdminRemovetextLabel5 = New-Object System.Windows.Forms.Label
        $AdminRemovetextLabel5.Location = '20,150'
        $AdminRemovetextLabel5.size = '330,40'
        $AdminRemovetextLabel5.Text = 'Will have the following actioned on their account:'
        $AdminRemoveSelectedMailBoxNametextLabel6 = New-Object System.Windows.Forms.Label
        $AdminRemoveSelectedMailBoxNametextLabel6.Location = '20,80'
        $AdminRemoveSelectedMailBoxNametextLabel6.Size = '350,40'
        $AdminRemoveSelectedMailBoxNametextLabel6.ForeColor = 'Blue'
        $AdminRemoveBox3 = New-Object System.Windows.Forms.GroupBox
        $AdminRemoveBox3.Location = '40,345'
        $AdminRemoveBox3.size = '650,30'
        $AdminRemoveBox3.text = '3. Click Ok to Process User or Exit:'
        $AdminRemoveBox3.button
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,385'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $CancelButton = new-object System.Windows.Forms.Button
        $CancelButton.Location = '470,385'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $AdminRemoveForm.Close()
                $AdminRemoveForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
        $AdminRemoveForm.Controls.AddRange(@($AdminRemoveBox1, $AdminRemoveBox2, $AdminRemoveBox3, $OKButton, $CancelButton))
        $AdminRemoveBox1.Controls.AddRange(@($AdminRemovetextLabel1, $AdminRemoveMBNameComboBox1))
        $AdminRemoveBox2.Controls.AddRange(@($AdminRemovetextLabel3, $AdminRemovetextLabel4, $AdminRemovetextLabel5, $AdminRemoveSelectedMailBoxNametextLabel6))
        $AdminRemoveForm.Add_Shown( { $AdminRemoveForm.Activate() })    
        $dialogResult = $AdminRemoveForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            #  CHECK - Don't accept no User selection   # 
            if ($AdminRemoveSelectedMailBoxNametextLabel6.Text -eq '') {
                [System.Windows.Forms.MessageBox]::Show("You need to select an Admin!`n`nTrying to enter blank fields is never a good idea.`n`nProcessing cannot continue.", $Title,
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $AdminRemoveForm.Close()
                $AdminRemoveForm.Dispose()
                break
            }
            $DateDel = (Get-Date).AddMonths(1).ToString('dd MMM yyyy') 
            # CHECK -  User has email address   #
            $Sam = Get-ADUser -filter { DisplayName -eq $AdminRemoveSelectedMailBoxNametextLabel6.Text } -Properties * | Select-Object SamAccountName | Select-object -ExpandProperty SamAccountName
            write-output $Sam
            try {
                Get-ADUser $Sam -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$($_.Description) YY DELETE after DATE - $Datedel" }
                Get-ADUser $Sam -Properties ProfilePath | ForEach-Object { Set-ADUser $_ -ProfilePath $Null }
                Get-ADUser $Sam -Properties HomeDirectory | ForEach-Object { Set-ADUser $_ -HomeDirectory $Null }
                Get-ADUser $Sam -Properties HomeDrive | ForEach-Object { Set-ADUser $_ -HomeDrive $Null } 
                Write-Verbose " Labelling Admin AD account to be deleted in 1 month and clearing Profile and P drive paths" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG labelling the Admin AD account to be deleted in 1 month !!!.`n`nPlease contact the Systems Integration Team with the details.", $Title,
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $AdminRemoveForm.Close()
                Return AdminRemove
            }
            Disable-ADAccount -Identity $Sam
            Write-Verbose "Disabling AD Account" -Verbose
            $poplabel = "Moving Admin account to the `n`nSCTS/User Accounts/Z-Disabled_Leavers OU."
            PopupForm
            try {
                Get-ADUser $Sam | Move-ADObject -targetpath 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
                Write-Verbose "moving Admin AD account to Z-Disabled_Leavers OU" -Verbose
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG moving the Admin account to the Z-Disabled_Leavers OU !!!.`n`nPlease contact the Systems Integration Team with the details.", $Title,
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $AdminRemoveForm.Close()
                Return AdminRemove 
            }
        }   
    }
    AdminRemove
}
