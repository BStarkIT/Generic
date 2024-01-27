param
(
    [Parameter(Mandatory)]$User
)
Set-Mailbox -Identity aberdeencriminalteam@scotcourts.gov.uk -GrantSendOnBehalfTo @{add=$User}
