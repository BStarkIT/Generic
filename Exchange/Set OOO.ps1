param
(
    [Parameter(Mandatory)]$User
)
$Reply = "<html><head></head><body><p>Having regard to the effects of Covid-19 pandemic.....</br>Apologies for any inconvenience.</br>Thank you</p></body></html>"

Set-MailboxAutoReplyConfiguration -Identity $User -AutoReplyState Enabled -InternalMessage $Reply -ExternalMessage $Reply