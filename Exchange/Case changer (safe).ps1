<#
.SYNOPSIS
This PowerShell script is to Change all email addresses to lowercase, without changing X400 or X500 values.
This is the SAFE version of the file, against the test OU & with write-output rather than the change AD values.

.NOTES
Script written by Brian Stark 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$userou = 'ou=soe users 2.6,ou=scts users,ou=user accounts,ou=scts,DC=scotcourts,DC=local'
$users = Get-ADUser -Filter * -SearchBase $userou -Properties ProxyAddresses, EmailAddress, SamAccountName


ForEach ($user in $users) {
    $Proxylist = @()
    foreach ($address in $user.ProxyAddresses) {
        if ($address -clike "SMTP:*") {
            $Name = $address.Split("@")[0]
            $Domain = $address.Split("@")[-1]
            $New = $Domain.ToLower()
            $Primary = $Name + "@" + $New
        }
        elseif ($address -like "X*") {
            $Proxylist += $address
        }
        else {
            $proxy = $address.ToLower()
            $Proxylist += $proxy
        }
    }
    #Set-ADUser -identity $User.SamAccountName -replace @{proxyAddresses = ($Primary) }
    foreach ($Mail in $Proxylist) {
        Write-Output $Mail
        #Set-ADUser -identity $User.SamAccountName -add @{proxyAddresses = ($Mail) }
    }
    $Email = $Primary.Split(":")[-1]
    #Set-ADUser -identity $User.SamAccountName -replace EmailAddress $Email
    Write-Output $Email
    $Name = $User.SamAccountName
    Write-Output "User $Name updated"
}