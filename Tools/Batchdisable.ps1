
foreach ($User in (Import-Csv -Path C:\PS\unused3.csv | Select-Object -ExpandProperty SAM )){
    $MailboxCheck = Get-ADUser $User -Properties * | Select-Object EmailAddress | Select-Object -ExpandProperty EmailAddress
    Write-Host "$User selected for deletion."
    $DateDel = (Get-Date).AddDays(180).ToString('dd MMM yyyy') 
    Write-Host "deletion date $DateDel"
    $Usermembership = get-aduser -identity $User -property MemberOf |
    Foreach-Object { $_.MemberOf | Get-AdGroup | Select-Object Name, SamaccountName | Select-object -ExpandProperty SamAccountName }
    $pdrive = test-path \\scotcourts.local\home\P\$User
    try {
        If (($null -ne $Usermembership) -and ($pdrive -eq $true)) {                   
            $Usermembership | out-file \\scotcourts.local\home\P\$User\UserMembershipBackup.csv
            Write-Host "Backed up Users Security group and Distribution List membership to $User P drive." 
        }
    }
    catch {
        Write-Host "Error - Cant Users Security group and Distribution List membership to $User P drive." 
    }
    Write-Host "Renaming $User P drive folders."
    try {
        if (test-path \\scotcourts.local\Home\P\$User) {
            Get-Item \\scotcourts.local\home\P\$User | Rename-Item -NewName { $_.Name -replace "$", " xx DELETE after DATE - $Datedel" }
            Write-Host "Renamed $User P drive folder on SAUFS01" 
        }
    }
    catch {
        Write-Host "Error - Cannot Renamed $User P drive folder on SAUFS01" 
    }
    If ($null -eq $MailboxCheck) {
        Write-Host "The selected User does not have a mailbox SAM - $User Processing cannot continue."
    }
    else {
        Write-Host "Hiding email address $MailboxCheck from global address list`n`nand`n`nSetting users mailbox to not receive emails"
        try {
            $email = get-mailbox $MailboxCheck | Select-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySMTPAddress
            if ($null -eq $email) {
                $O365Email = get-remotemailbox $MailboxCheck | Select-Object PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySMTPAddress
                set-remotemailbox -identity $O365Email -AcceptMessagesOnlyFrom 'helpdesk' -HiddenFromAddressListsEnabled $true
            }
            else {
                Get-Mailbox $MailboxCheck | Set-Mailbox -AcceptMessagesOnlyFrom 'helpdesk' -HiddenFromAddressListsEnabled $true
            }
            Write-Host "Setting $MailboxCheck mailbox not to receive emails and hiding from address list" 
        }
        catch {
            Write-Host "Error - Cannot set $MailboxCheck mailbox not to receive emails" 
        }
    }
    Write-Host "Disabling Users account."
    try {
        Disable-ADAccount -Identity $User
        Write-Host "Disabling AD Account $User" 
    }
    catch {
        Write-Host "Error - Cannot disable AD Account $User"
        break
    }
    Write-Host "Editing users AD account: Labelling to be deleted in 6 months. Clearing P drive path."
    try {
        Get-ADUser $User -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$($_.Description) xx DELETE after DATE -- $Datedel" }
        Write-Host " Labelling $User AD account to be deleted in 1 month and clearing P drive paths" 
    }
    catch {
        Write-Host "Error when Labelling $User AD account to be deleted in 1 month and clearing P drive paths"
        break
    }
    Write-Host "Moving user account to the `n`nSCTS/User Accounts/Z-Disabled_Leavers OU."
    try {
        Get-ADUser $User | Move-ADObject -targetpath 'OU=CRN354-22,OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
        Write-Host "moving $User AD account to Z-Disabled_Leavers OU" 
    }
    catch {
        Write-Host "Error moving $User AD account to Z-Disabled_Leavers OU" 
        break
    }
    Write-host "Checking and removing`n`nany Users Shared Mailbox permissions"
    $SharedMailboxes = Get-ADUser -Identity $User -Properties msExchDelegateListBL | Select-Object -ExpandProperty msExchDelegateListBL
    $SharedMailboxes -replace '^CN=|,.*$'
    foreach ($SharedMailbox in $SharedMailboxes) {
        $Shared = get-mailbox -identity $SharedMailbox | select-object PrimarySMTPAddress | Select-Object -ExpandProperty PrimarySMTPAddress
        if ($null = $Shared) {
            $O365Shared = get-mailbox -identity $SharedMailbox | select-object PrimarySMTPAddress | Select-Object -ExpandProperty PrimarySMTPAddress
            Write-Host "Please remove $User from $O365Shared via the O365 exchange control pannel"
        }
        else {
            $O365Shared = $null
            Remove-MailboxPermission $Shared -User $User -AccessRights FullAccess -confirm:$false 
            Set-Mailbox $Shared -GrantSendOnBehalfTo @{remove = "$User" }
            Write-Host "Removed $User Shared Mailbox permissions" 
        }
    }
    Write-Host "Checking and removing User from`n`nSecurity groups`n`nand Distribution Lists."
    try {
        If ($null -ne $Usermembership) {                    
            $Usermembership | ForEach-Object {
                Remove-ADGroupMember -Identity $_ -Member $User -confirm:$false
                Write-Host "Removed $User from groups" }
        }
    }
    catch {
        Write-Host "Error Removing $User from groups"
        break
    }
    $RBACUser = "$User" + "_a"
    try {
        $RBACCatch = Get-ADUser -Identity $RBACUser -Properties * | Select-Object SamAccountName | Select-Object -ExpandProperty SamAccountName
        Get-ADUser $RBACCatch -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$($_.Description) xx DELETE after DATE - $Datedel" }
        Get-ADUser $RBACCatch | Move-ADObject -targetpath 'OU=Z-Disabled_Leavers,OU=User Accounts,OU=SCTS,DC=scotcourts,DC=local'
        Write-Host "moving $RBACCatch AD account to Z-Disabled_Leavers OU" 
    }
    catch {
        Write-Output "no RBAC account"
    }
    if ($Null -eq $O365Shared) {
        Write-Host "Complete"
    }
    else {
        Write-Host "Complete but O365 user"
    }
}