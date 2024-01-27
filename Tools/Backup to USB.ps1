<#
.SYNOPSIS
This PowerShell script is to backup user profile folder to USB drive, if 2 drives are present, prompt to select destination drive.

.NOTES
Script written by Brian Stark
Date:
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
$List = gwmi win32_diskdrive | Where-Object { $_.interfacetype -eq "USB" } | ForEach-Object { gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition" } | ForEach-Object { gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition" } | ForEach-Object { $_.deviceid }
MainForm
function Backup {
    Write-Output "backing up to $Drive"
    Robocopy "$env:UserProfile" $Drive /XD AppData .* /XF NTUSER.DAT *.dat.LOG* /MT:32 /E /Z 
}
Function MainForm {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    # Set the details size of your form
    $ManForm = New-Object System.Windows.Forms.Form
    $ManForm.width = 350
    $ManForm.height = 200
    $ManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $ManForm.MinimizeBox = $False
    $ManForm.MaximizeBox = $False
    $ManForm.FormBorderStyle = 'Fixed3D'
    $ManForm.Text = 'Destination selection'
    $ManFormNameComboBox1 = New-Object System.Windows.Forms.ComboBox
    $ManFormNameComboBox1.Location = '25,35'
    $ManFormNameComboBox1.Size = '50, 50'
    $ManFormNameComboBox1.AutoCompleteMode = 'Suggest'
    $ManFormNameComboBox1.AutoCompleteSource = 'ListItems'
    $ManFormNameComboBox1.Sorted = $true;
    $ManFormNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $ManFormNameComboBox1.DataSource = $List
    $ManFormNameComboBox1.add_SelectedIndexChanged( { $ManFormNameComboBox1.Text = "$($ManFormNameComboBox1.SelectedItem.ToString())" })
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '100,100'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '200,100'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = 'Exit'
    $CancelButton.add_Click( {
            $ManForm.Close()
            $ManForm.Dispose()
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    $ManForm.Controls.AddRange(@($ManFormNameComboBox1, $OKButton, $CancelButton))
    $ManForm.AcceptButton = $OKButton
    $ManForm.CancelButton = $CancelButton
    $ManForm.Add_Shown( { $ManForm.Activate() })    
    $Result = $ManForm.ShowDialog()
    if ($Result -eq 'OK') {
        if ($AddACLFullSelectedUserNametextLabel4.Text -eq '') {
            [System.Windows.Forms.MessageBox]::Show("You need to select a USB drive", 'Backup to usb.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $AddACLFullForm.Close()
            $AddACLFullForm.Dispose()
            break
        }
        else {
            $Drive = $ManFormNameComboBox1.Text
            Backup
        }
    }
}
