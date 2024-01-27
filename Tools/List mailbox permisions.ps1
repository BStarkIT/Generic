param
(
    [Parameter(Mandatory)]$Email
)
<<<<<<< HEAD
$Users = Get-MailboxPermission –identity $Email | Select-Object User
foreach ($User in $Users
Get-ADUser $User )
=======
 Get-MailboxPermission –identity $Email
>>>>>>> 3a7ecd6... .
