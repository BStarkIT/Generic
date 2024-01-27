param
(
    [Parameter(Mandatory)]$User
)
Get-ADUser $User -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet