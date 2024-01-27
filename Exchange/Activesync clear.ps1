param
(
    [Parameter(Mandatory)]$User
)
Set-CASMailbox -Identity $User -ActiveSyncEnabled $True
Get-ActiveSyncDevice -Mailbox $User | Remove-ActiveSyncDevice
