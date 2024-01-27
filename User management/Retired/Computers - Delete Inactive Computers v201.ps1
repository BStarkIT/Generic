# Delete Inactive Computers v2.01
# Author        John Mckay
# Date          18/10/2018
# Version       2.01
# Purpose       To delete disabled computer accounts marked for deletion over 30 days old.
# Useage        helpdesk staff have a monthly task on servicedesk to run a shorcut to Delete Inactive Computers.
#               Script searches the courts\computers\Z-Disabled_InactiveFor60Days ou for computers that have a delete date older than 30 days AND are disabled and deletes computer account.
#               An email is sent to helpdesk, infrastructure, governance and dst with details of the computers that have been deleted.
#
# Changes       
#
# Script function: This script performs the following on a computer account.
$message = "
This script searches for computers in the Z-Disabled_InactiveFor60Days ou.

If the 'delete after date' is over 30 days the computer will be processed as below.

AD -  computer account will be deleted.

Email - a list of deleted computers will be emailed to Helpdesk, Infrastructure & DST."
#
# Start of script:
#
##  Show Start Message:   ###
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null 
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", 'Delete Inactive Computers v2.01', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {Exit} 
else {    
    ## ## create the pop up information form   ###
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        # create form
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = 'Delete Inactive Computers.'
        $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $PopForm.Size = New-Object System.Drawing.Size(420, 200)
        # Add Label
        $PopUpLabel = New-Object System.Windows.Forms.Label
        $PopUpLabel.Location = '80, 40' 
        $PopUpLabel.Size = '300, 80'
        $PopUpLabel.Text = $poplabel
        $PopForm.Controls.Add($PopUpLabel)
        # Show the form
        $PopForm.Show()| Out-Null
        # wait 2 seconds
        Start-Sleep -Seconds 2
        # close form
        $PopForm.Close() | Out-Null
    }
    ## create the pop up information form complete
    #
    ## create function to delete multiple computers   ###
    Function DeleteMultipleComputers {
        ## get list of computers with name and description to display in form   ###
        $ComputersListToDisplayInForm = foreach ($computer in $ComputersOlderThan1MonthList) {$($computer.Name) + " - " + $($computer.Description)}
        ## create main form   ###
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $DeletingComputersForm = New-Object System.Windows.Forms.Form
        $DeletingComputersForm.width = 745
        $DeletingComputersForm.height = 495
        $DeletingComputersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $DeletingComputersForm.Controlbox = $false
        $DeletingComputersForm.Icon = $Icon
        $DeletingComputersForm.FormBorderStyle = 'Fixed3D'
        $DeletingComputersForm.Text = 'Delete Inactive Computers v2.01'
        $DeletingComputersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $ComputerBox1 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox1.Location = '40,20'
        $ComputerBox1.size = '650,90'
        $ComputerBox1.text = '1. Check: This deletes computer accounts:'
        ### Create group 1 box text labels. ###
        $ComputertextLabel1 = New-Object System.Windows.Forms.Label
        $ComputertextLabel1.Location = '20,25'
        $ComputertextLabel1.size = '600,30'
        $ComputertextLabel1.Text = 'Check the computers in the list below to ensure thay are all marked to be deleted.' 
        $ComputertextLabel2 = New-Object System.Windows.Forms.Label
        $ComputertextLabel2.Location = '20,55'
        $ComputertextLabel2.size = '600,30'
        $ComputertextLabel2.Text = 'IF THERE ARE ANY COMPUTERS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
        ### Create group 2 box in form. ###
        $ComputerBox2 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox2.Location = '40,120'
        $ComputerBox2.size = '650,225'
        $ComputerBox2.text = '2. Check the Computers below are labelled to be Deleted'
        # Create group 2 box text labels.
        $ComputertextLabel4 = New-Object System.Windows.Forms.ListBox    
        $ComputertextLabel4.Location = '40,20'
        $ComputertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $ComputertextLabel4.size = '570,170'
        $ComputertextLabel4.Datasource = $ComputersListToDisplayInForm 
        ### Create group 3 box in form. ###
        $ComputerBox3 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox3.Location = '40,355'
        $ComputerBox3.size = '650,30'
        $ComputerBox3.text = '3. Click Ok to DELETE Computer Accounts or Exit:'
        $ComputerBox3.button
        ### Add an OK button ###
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,395'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '470,395'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $DeletingComputersForm.Close()
                $DeletingComputersForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        ### Add all the Form controls ### 
        $DeletingComputersForm.Controls.AddRange(@($ComputerBox1, $ComputerBox2, $ComputerBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $ComputerBox1.Controls.AddRange(@($ComputertextLabel1, $ComputertextLabel2))
        $ComputerBox2.Controls.AddRange(@($ComputertextLabel4))
        #### Activate the form ###
        $DeletingComputersForm.Add_Shown( {$DeletingComputersForm.Activate()})    
        #### Get the results from the button click ###
        $dialogResult = $DeletingComputersForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteMultipleComputers form and starting to Remove computer AD Account" -Verbose
            DeleteComputerADAccount
        }
    } 
    ## function to delete multiple computers complete   ###
    #
    ## function to delete one computer   ###
    Function DeleteOneComputer {
        ## get list of computers with name and description to display in form   ###
        $ComputersListToDisplayInForm = foreach ($computer in $ComputersOlderThan1MonthList) {$($computer.Name) + " - " + $($computer.Description)}
        ## create main form   ###
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
        ### Set the details of the form. ###
        $DeletingComputersForm = New-Object System.Windows.Forms.Form
        $DeletingComputersForm.width = 745
        $DeletingComputersForm.height = 495
        $DeletingComputersForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
        $DeletingComputersForm.Controlbox = $false
        $DeletingComputersForm.Icon = $Icon
        $DeletingComputersForm.FormBorderStyle = 'Fixed3D'
        $DeletingComputersForm.Text = 'Delete Inactive Computers v2.01'
        $DeletingComputersForm.Font = New-Object System.Drawing.Font('Ariel', 10)
        ### Create group 1 box in form. ####
        $ComputerBox1 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox1.Location = '40,20'
        $ComputerBox1.size = '650,90'
        $ComputerBox1.text = '1. Check: This deletes computer accounts:'
        ### Create group 1 box text labels. ###
        $ComputertextLabel1 = New-Object System.Windows.Forms.Label
        $ComputertextLabel1.Location = '20,25'
        $ComputertextLabel1.size = '600,30'
        $ComputertextLabel1.Text = 'Check the computer below to ensure it is marked to be deleted.' 
        $ComputertextLabel2 = New-Object System.Windows.Forms.Label
        $ComputertextLabel2.Location = '20,55'
        $ComputertextLabel2.size = '600,30'
        $ComputertextLabel2.Text = 'IF THE COMPUTER IS NOT MARKED TO BE DELETED DO NOT PROCEED.' 
        ### Create group 2 box in form. ###
        $ComputerBox2 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox2.Location = '40,120'
        $ComputerBox2.size = '650,225'
        $ComputerBox2.text = '2. Check the Computer below is labelled to be Deleted'
        # Create group 2 box text labels.
        $ComputertextLabel4 = New-Object System.Windows.Forms.Label
        $ComputertextLabel4.Location = '20,45'
        $ComputertextLabel4.size = '570,170'
        $ComputertextLabel4.Font = New-Object System.Drawing.Font('Ariel', 8)
        $ComputertextLabel4.Text = "$ComputersListToDisplayInForm"
        ### Create group 3 box in form. ###
        $ComputerBox3 = New-Object System.Windows.Forms.GroupBox
        $ComputerBox3.Location = '40,355'
        $ComputerBox3.size = '650,30'
        $ComputerBox3.text = '3. Click Ok to DELETE Computer Account or Exit:'
        $ComputerBox3.button
        ### Add an OK button ###
        $OKButton = new-object System.Windows.Forms.Button
        $OKButton.Location = '590,395'
        $OKButton.Size = '100,40'          
        $OKButton.Text = 'Ok'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        ### Add a cancel button ###
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '470,395'
        $CancelButton.Size = '100,40'
        $CancelButton.Text = 'Exit'
        $CancelButton.add_Click( {
                $DeletingComputersForm.Close()
                $DeletingComputersForm.Dispose()
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel})
        ### Add all the Form controls ### 
        $DeletingComputersForm.Controls.AddRange(@($ComputerBox1, $ComputerBox2, $ComputerBox3, $OKButton, $CancelButton))
        #### Add all the GroupBox controls ###
        $ComputerBox1.Controls.AddRange(@($ComputertextLabel1, $ComputertextLabel2))
        $ComputerBox2.Controls.AddRange(@($ComputertextLabel4))
        #### Activate the form ###
        $DeletingComputersForm.Add_Shown( {$DeletingComputersForm.Activate()})    
        #### Get the results from the button click ###
        $dialogResult = $DeletingComputersForm.ShowDialog()
        # If the OK button is selected
        if ($dialogResult -eq 'OK') {
            Write-Verbose " Accepted DeleteOneComputer form and starting to Remove computer AD Account" -Verbose
            DeleteComputerADAccount
        }
    }
    ## function to delete one computer complete   ###
    #  
    ## create the Delete AD account function   ###
    Function DeleteComputerADAccount {
        ## create a list to add the deleted user names to   ###
        $DeleteADComputerList = ""
        try {
            ForEach ($ComputerToDelete in $ComputersOlderThan1MonthList) {
                ## delete computer AD account   ###
                Remove-ADComputer -Identity $ComputerToDelete.Name -Confirm:$false
                ## add computer name to list   ###
                $DelADComputerList = $($ComputerToDelete.Name)
                $DeleteADComputerList += "$DelADComputerList`r`n"
                Write-Verbose "The AD account for computer $($ComputerToDelete.Name) has been deleted." -Verbose 
                $poplabel = "Deleting the AD account for computer`n`n$($ComputerToDelete.Name)."
                PopupForm
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG deleting the User`n`n$($ComputerToDelete.Name)`n`nAD account !!!.`n`nPlease contact the Infrastructure Team with the details.", 'Delete Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            break
        }
        ## Details of deleted AD user accounts to include in email to helpdesk   ###
        $DeletedADListForMessage = @()
        $DeletedADListForMessage = $DeleteADComputerList
        $DeletedADListForMessageTxt = ''
        $DeletedADListForMessage | ForEach-Object { $DeletedADListForMessageTxt += $_ + "`n" }
        Write-Verbose "The Computers with a delete date older than 1 month have been deleted." -Verbose
        ##  send email to helpdesk, infrastructure, governance and dst   ###
        $message2 = "The computers below have been deleted from AD.`n`n$DeletedADListForMessageTxt"  
        $mailrecipients = "helpdesk@scotcourts.gov.uk", "dst@scotcourts.gov.uk", "itinfrastructure@scotcourts.gov.uk", "itgovernance@scotcourts.gov.uk"
        Send-MailMessage -To $mailrecipients -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: AD Computers - Deleted Inactive Computers $(Get-Date -format ('dd MMMM yyyy'))" -Body "$message2" -SmtpServer mail.scotcourts.local
        ##  Message complete message   ###
        [System.Windows.Forms.MessageBox]::Show("The Delete Inactive Computers process is complete.`n`nA list of deleted computers will be emailed.", 'Delete Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        break
    } 
    ## Delete AD account function complete   ### 
    #
    ### Set Delete Date to -1 month   ###
    $DatetoDel = (Get-Date).AddMonths(-1).ToString('dd MMMM yyyy')
    ## get list of computers in the Z-Disabled_InactiveFor60Days ou that are marked to be deleted and disabled   ###
    $ComputersList = Get-ADComputer -Filter * -SearchBase "OU=Z-Disabled_InactiveFor60Days,OU=Computers,ou=courts,DC=Scotcourts,DC=Local" -Properties Name, Description, Enabled |
        Select-Object Name, Description, Enabled | Where-Object {$_.Description -match "xx DELETE after DATE" -and $_.Enabled -eq $false}
    ## get users older than 1 month   ###
    $ComputersOlderThan1MonthList = ForEach ($ComputerToDelete in $ComputersList) {
        ## get date on users AD description   ###
        $ComputerInfo = Get-ADComputer -identity $ComputerToDelete.Name -Properties * |
            Select-Object Name, Description | select-object Name, @{n = 'DeleteComputerOnDate'; e = {$_.Description -replace '^.*-'}} 
        $ComputerDelDate = [System.DateTime]$ComputerInfo.DeleteComputerOnDate
        ## check if computer delete date is over 1 month   ###
        If ($ComputerDelDate -lt $DateToDel) {
            Get-AdComputer $ComputerToDelete.Name -Properties * | select-object Name, Description
            Write-Verbose "The AD account for computer $($ComputerToDelete.Name) is past its delete date." -Verbose
        }
    }
    ## if no computers to delete show message and exit   ###
    if ($ComputersOlderThan1MonthList -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("There are currently no Inactive Computers to delete.`n`nThe Delete Inactive Computers process is complete.", 'Delete Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        break
    }
    if (@($ComputersOlderThan1MonthList).Count -eq 1) {
        DeleteOneComputer
    }
    Else {
        DeleteMultipleComputers
    }
}
 