$Shared = 'dundeehcjury@scotcourts.gov.uk'
$user = 'mmcfarlane@scotcourts.gov.uk'

ForEach ($Share in $Shared) {
    Add-MailboxPermission -Identity $Share -User $user -AccessRights FullAccess -InheritanceType All    
    Set-Mailbox -Identity $Share -GrantSendOnBehalfTo @{add=$User}
}
