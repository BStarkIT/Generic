
$userou = 'OU=AzureAD,OU=User Accounts (Testing),OU=SCTS,DC=scotcourts,DC=local'
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