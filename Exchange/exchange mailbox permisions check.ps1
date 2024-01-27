param
(
    [Parameter(Mandatory)]$Email
)
get-MailboxPermission -Identity $Email
