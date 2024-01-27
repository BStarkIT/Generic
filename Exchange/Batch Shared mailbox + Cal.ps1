$Shares = 'hpcadmin@scotcourtstribunals.gov.uk'
$user = 'clee@scotcourts.gov.uk'

ForEach ($Share in $Shares) {
    #Remove-MailboxPermission -Identity $share -user $user -AccessRights FullAccess -InheritanceType All -Confirm:$false
    Add-MailboxPermission -Identity $Share -User $user -AccessRights FullAccess -InheritanceType All
    Set-Mailbox -Identity $Share -GrantSendOnBehalfTo @{add=$User}
    #add-MailboxfolderPermission -Identity ${Share}:\calendar -User $user -AccessRights editor
    #remove-MailboxfolderPermission -Identity ${Share}:\calendar -User $user -Confirm:$false
    #Set-CalendarProcessing -Identity $Share -AutomateProcessing AutoAccept
    #Add-MailboxFolderPermission -Identity $Share:\Calendar -User $User -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems # for private marked email in a shared mailbox.
}
#set-mailbox $Share -MessageCopyForSendOnBehalfEnabled $True
#set-mailbox $Share -MessageCopyForSentAsEnabled $True