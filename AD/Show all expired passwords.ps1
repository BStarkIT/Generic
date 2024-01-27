##
## Simple script to list accounts with expired passwords
##
Get-ADUser -SearchBase "DC=ad,DC=realise,DC=com" -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, Office | Where-Object-Object { $_.Enabled -eq "True" } | Where-Object { $_.PasswordNeverExpires -eq $false } | Where-Object { $_.passwordexpired -eq $true }
