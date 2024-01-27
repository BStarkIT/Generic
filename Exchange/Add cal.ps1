param
(
    [Parameter(Mandatory)]$User,
    [Parameter(Mandatory)]$Shared
)
add-MailboxfolderPermission -Identity ${Shared}:\calendar -User $User -AccessRights Editor  
