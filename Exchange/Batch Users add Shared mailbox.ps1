$Shared = 'CyberSecurity@scotcourts.gov.uk'
$users = 'tsolowczuk', 'schristie2', 'jborgennielsen'
ForEach ($user in $users) {
    Add-MailboxPermission -Identity $Shared -User $user -AccessRights FullAccess -InheritanceType All    
    Set-Mailbox -Identity $Shared -GrantSendOnBehalfTo @{add = $User }
}