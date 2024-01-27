# Disable Inactive Computers v2.01
# Author        John Mckay
# Date          17/10/2018
# Version       2.01
# Purpose       To disable computers in AD that haven't been logged onto for 60 days or more.
# Useage        helpdesk staff have a weekly task on servicedesk to run a shorcut to Disable Inactive Computers.
#               Script disables computers & moves them to the  courts\computers\Z-Disabled_InactiveFor60Days OU.
#               An email is sent to helpdesk,dst & infrastructure with details of the disabled computers.
#
# Changes       
#
# Script function: This script performs the following on computer accounts.
$message = "
This script searches AD for computers that have not logged on for more than 60 days.

Any computer found will be processed as below.

AD - Disable computer account.

AD - Move computer to the Z-Disabled_InactiveFor60Days ou.

AD - Computer description labelled as Disabled and to be
         deleted in 1 month.

Email - A list of disabled computers will be emailed to Helpdesk,
             Infrastructure & DST."
#
# Start of script:
##  Show Start Message:   ###
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$StartMessage = [System.Windows.Forms.MessageBox]::Show("$message", 'Computers - Disable Inactive Computers v2.01', [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
if ($StartMessage -eq 'Cancel') {Exit} 
else {
    ##  Set Exclusion list of OU's not to be included   ###
    $OUExclusionList = 'OU='
    $OUExclusionList += @(
        'CitrixServers,DC=scotcourts,DC=local',
        'Domain Servers,DC=scotcourts,DC=local',
        'Domain Controllers,DC=scotcourts,DC=local',
        'Exchange Servers,DC=scotcourts,DC=local',
        'Servers Non Windows,DC=scotcourts,DC=local',
        'TribunalServers,OU=TRIBUNALS,DC=scotcourts,DC=local',
        'DARServers,OU=Computers By Location,DC=scotcourts,DC=local',
        'DAR2,OU=Computers By Location,DC=scotcourts,DC=local',
        'DAR Computers,OU=Computers,OU=Courts,DC=scotcourts,DC=local',
        'DARServersNLE,OU=Servers,OU=Computers,OU=Testing OU,DC=scotcourts,DC=local',
        'Servers,OU=Courts,DC=scotcourts,DC=local',
        'Z-Disabled_InactiveFor60Days,OU=Computers,OU=Courts,DC=scotcourts,DC=local'
    ) -join '|OU='
    #
    ## create the pop up information form   ###
    Function PopUpForm {
        Add-Type -AssemblyName System.Windows.Forms    
        # create form
        $PopForm = New-Object System.Windows.Forms.Form
        $PopForm.Text = 'Disable Inactive Computers not logged on for 60 days or more.'
        $PopForm.Size = New-Object System.Drawing.Size(420, 200)
        $PopForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
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
    ## create the pop up information form complete   ###
    #
    Function Process60DaysComputers {
        ##  set search date to 60 days ago  ###
        $date = [DateTime]::Today.AddDays(-60)
        ## set Delete Date for 1 month ahead  ###
        $DateDel = (Get-Date).AddMonths(1).ToString('dd MMM yyyy') 
        #
        ##  Get computers that haven't logged on for 60 days or over except those in exception list  ###
        $60DaysComputers = Get-ADComputer -Filter  ‘LastlogonDate -le $date’ -Properties * | Where-Object {$_.distinguishedname -notmatch $OUExclusionList} |
        Select-Object @{Name = "ComputerName"; Expression = {$_.Name}}, @{Name = "LastlogonDate"; Expression = {$_.LastlogonDate}}, @{Name = "OU"; Expression = {$_.DistinguishedName -replace "CN=$($_.Name),", ""}}, @{Name = "Description"; Expression = {$_.Description}}|
        ## if no inactive computers display message and exit   ###
        if ($60DaysComputers -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("There are currently no Inactive Computers in AD.`n`nThe Disable Inactive Computers process is complete.", 'Disable Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            break            
        }
        Else {
            ## create a list to add the disabled computers to   ###
            $DisabledComputerList = ""
            ForEach ($Computer in $60DaysComputers) {
                ##  disable computer AD account   ##
                try {
                    Set-ADComputer $Computer.ComputerName -Enabled $false
                    ## add computer name to list   ###
                    $DisADComputerList = "$($Computer.ComputerName)"
                    $DisabledComputerList += "$DisADComputerList`r`n"
                    Write-Verbose "Disabling computer $($Computer.ComputerName) AD Account" -Verbose
                    $poplabel = "Disabling Computer $($Computer.ComputerName) AD account."
                    PopupForm
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG disabling computer $($Computer.ComputerName) account !!!.`n`nPlease contact the Infrastructure Team with the details.", 'Disable Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    break 
                }
                ##  move computer account to Z-Disabled_InactiveFor60Days OU   ###
                try {
                    get-adcomputer $Computer.ComputerName | Move-ADObject -TargetPath 'OU=Z-Disabled_InactiveFor60Days, OU=computers, OU=courts, DC=scotcourts, DC=local'
                    Write-Verbose "Moving computer $($Computer.ComputerName) AD account to Z-Disabled_InactiveFor60Days OU" -Verbose
                    $poplabel = "Moving computer $($Computer.ComputerName) account to the `n`nZ-Disabled_InactiveFor60Days OU OU."
                    PopupForm
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG moving the computer $($Computer.ComputerName) account to the Z-Disabled_InactiveFor60Days !!!.`n`nPlease contact the Infrastructure with the details.", 'Disable Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    break  
                }
                ##  edit computer AD description field to delete account in 1 month   ###
                $poplabel = "Editing computer $($Computer.ComputerName) AD description:`n`nLabelling to be deleted in 1 month."
                PopupForm
                try {
                    Set-ADComputer -Identity $Computer.ComputerName -Description "DISABLED - DO NOT ENABLE WITHOUT AUTHORISATION - xx DELETE after DATE - $Datedel"
                    Write-Verbose " Labelling computer $($Computer.ComputerName) AD account to be deleted in 1 month" -Verbose
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Something has gone WRONG labelling the computer $($Computer.ComputerName) AD description to be deleted in 1 month !!!.`n`nPlease contact the Infrastructure Team with the details.", 'Disable Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    break 
                }
            }
            ##  Details of disabled computer accounts to include in email   ###
            $DisabledADListForMessage = @()
            $DisabledADListForMessage = $DisabledComputerList
            $DisabledADListForMessageTxt = ''
            $DisabledADListForMessage | ForEach-Object { $DisabledADListForMessageTxt += $_ + "`n" }
            #
            ##  send email to helpdesk, infrastructure and dst   ###
            $message2 = "The following computers haven't been logged onto in 60 days and have been disabled and`n`nmoved into the courts\computers\Z-Disabled_InactiveFor60Days ou.`n`nDO NOT ENABLE ANY OF THESE COMPUTERS IN AD`n`nWITHOUT AUTHORISATION FROM IT SENIOR MANAGEMENT.`n`n$DisabledADListForMessageTxt"  
            $mailrecipients = "helpdesk@scotcourts.gov.uk", "dst@scotcourts.gov.uk", "itinfrastructure@scotcourts.gov.uk"
            Send-MailMessage -To $mailrecipients -from $env:UserName@scotcourts.gov.uk -Subject "HDupdate: AD Computers - Disabled Computer accounts due to inactivity $(Get-Date -format ('dd MMMM yyyy'))" -Body "$message2" -SmtpServer mail.scotcourts.local
            ##  Message complete message   ###
            [System.Windows.Forms.MessageBox]::Show("The Disable Inactive Computers process is complete.`n`nA list of disabled computers will be emailed.", 'Disable Inactive Computers.', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            break
        }       
    } 
    Process60DaysComputers 
}        

