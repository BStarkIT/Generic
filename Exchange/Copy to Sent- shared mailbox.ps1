param
(
    [Parameter(Mandatory)]$Mailbox
)
set-mailbox $Mailbox -MessageCopyForSendOnBehalfEnabled $True
set-mailbox $Mailbox -MessageCopyForSentAsEnabled $True
