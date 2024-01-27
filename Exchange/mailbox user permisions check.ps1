param
(
    [Parameter(Mandatory)]$Email,
    [Parameter(Mandatory)]$User
)
get-MailboxPermission -Identity $Email  -User $User
