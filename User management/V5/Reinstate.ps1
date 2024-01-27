<#
.SYNOPSIS
This PowerShell script is to fetch computer information.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$Version = "1.00"
$Leavers = Get-ADUser -filter * -searchbase 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local' -Properties DisplayName | Select-Object Displayname | Select-Object -ExpandProperty DisplayName
$WinTitle = "re-instate user $Version"
$Icon = '\\scotcourts.local\data\CDiScripts\Scripts\Resources\Icons\EditUser.ico'
$UserName = $env:username
if ($UserName -notlike "*_a") {
    Write-Host "Must be run as Admin, Script run as $UserName"
    Pause
    
}
else {

function Activate {
    $newdescription = $UserDescription.Split("xx")[0]
    $Date = $UserDescription.Split("xx")[-1]
    $Path = $UserSamAccountName + " xx " + $Date
    $grouplist = Get-Content "\\scotcourts.local\home\P\$Path\UserMembershipBackup.csv"
    Rename-Item "\\scotcourts.local\home\P\$Path $UserSamAccountName"
    Enable-ADAccount -Identity $UserSamAccountName
    Set-ADUser -identity $UserSamAccountName -Description $newdescription
    Set-ADUser -identity $UserSamAccountName -Clear msExchHideFromAddressLists
    Set-ADUser -identity $UserSamAccountName -Clear authOrig
    foreach ($ADgroup in $grouplist) {
        Add-ADGroupMember -Identity $ADgroup -Members $UserSamAccountName
    }
    Get-ADUser $UserSamAccountName | Move-ADObject -targetpath 'OU=SOE Users 2.6,OU=SCTS Users,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
    Write-host "User account $UserSamAccountName has been reinstated"
    Pause
}
    function MainForm {
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        $reinstateManForm = New-Object System.Windows.Forms.Form
        $reinstateManForm.width = 550
        $reinstateManForm.height = 375
        $reinstateManForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $reinstateManForm.MinimizeBox = $False
        $reinstateManForm.MaximizeBox = $False
        $reinstateManForm.FormBorderStyle = 'Fixed3D'
        $reinstateManForm.Text = $WinTitle
        $reinstateManForm.Icon = $Icon
        $reinstateManForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        $Logo = [System.Drawing.Image]::Fromfile('\\saufs01\IT\Enterprise Team\Usermanagement\icons\SCTS.png')
        $pictureBox = New-Object Windows.Forms.PictureBox
        $pictureBox.Width = $Logo.Size.Width
        $pictureBox.Height = $Logo.Size.Height
        $pictureBox.Image = $Logo
        $reinstateManForm.controls.add($pictureBox)
        $reinstateManFormtext1 = New-Object System.Windows.Forms.Label
        $reinstateManFormtext1.Location = '20,120'
        $reinstateManFormtext1.size = '500,50'
        $reinstateManFormtext1.Text = "This script changes the following on a Users account: `n - AD account - reinstate & move."
        $EmailBox1 = New-Object System.Windows.Forms.GroupBox
        $EmailBox1.Location = '10,175'
        $EmailBox1.size = '500,75'
        $EmailBox1.text = '1. Select a UserName from the dropdown lists:'
        $reinstateManFormtextLabel1 = New-Object System.Windows.Forms.Label
        $reinstateManFormtextLabel1.Location = '20,40'
        $reinstateManFormtextLabel1.size = '100,20'
        $reinstateManFormtextLabel1.Text = 'UserName:' 
        $reinstateManFormNameComboBox1 = New-Object System.Windows.Forms.ComboBox
        $reinstateManFormNameComboBox1.Location = '125,35'
        $reinstateManFormNameComboBox1.Size = '350, 310'
        $reinstateManFormNameComboBox1.AutoCompleteMode = 'Suggest'
        $reinstateManFormNameComboBox1.AutoCompleteSource = 'ListItems'
        $reinstateManFormNameComboBox1.Sorted = $true;
        $reinstateManFormNameComboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
        $reinstateManFormNameComboBox1.DataSource = $Leavers
        $reinstateManFormNameComboBox1.add_SelectedIndexChanged( { $ChangeThisUser.Text = "$($reinstateManFormNameComboBox1.SelectedItem.ToString())" })
        $reinstateManFormtext2 = New-Object System.Windows.Forms.Label
        $reinstateManFormtext2.Location = '20,275'
        $reinstateManFormtext2.size = '75,150'
        $reinstateManFormtext2.Text = 'Reinstate:'
        $ChangeThisUser = New-Object System.Windows.Forms.Label
        $ChangeThisUser.Location = '100,275'
        $ChangeThisUser.Size = '200,50'
        $ChangeThisUser.ForeColor = 'Blue'
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
                $reinstateManForm.Close()
                $reinstateManForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
        $reinstateManForm.Controls.AddRange(@($EmailBox1, $ChangeThisUser, $reinstateManFormtext1, $reinstateManFormtext2, $OKButton, $CancelButton))
        $EmailBox1.Controls.AddRange(@($reinstateManFormtextLabel1, $reinstateManFormNameComboBox1))
        $reinstateManForm.AcceptButton = $OKButton
        $reinstateManForm.CancelButton = $CancelButton
        $reinstateManForm.Add_Shown( { $reinstateManForm.Activate() })    
        $Result = $reinstateManForm.ShowDialog()
        if ($Result -eq 'OK') {
            if ($ChangeThisUser.Text -eq '') {
                Write-Host "no user selected"
                [System.Windows.Forms.MessageBox]::Show("You need to select a Username !!!!!  Trying to enter blank fields is never a good idea.", 'Activator.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                Return MainForm
            }
            $SelectedUser = $ChangeThisUser.text
            $UserSamAccountName = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" | Select-Object -ExpandProperty 'SamAccountName'
            $UserDescription = Get-ADUser -Filter "Displayname -eq '$SelectedUser'" -Properties * | Select-Object -ExpandProperty description
            Write-Host "$UserSamAccountName Selected"
            Activate
        }
    }
    MainForm
}
