$un = read-Host 'Please enter username of person to Check:'
Get-ADUser $un -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet