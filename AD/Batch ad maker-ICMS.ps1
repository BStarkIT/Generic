$number = 1
$Name = "SCTSDev"
$OU = "OU=AzureAD,OU=User Accounts (Testing),OU=SCTS,DC=scotcourts,DC=local"
$Description = "ICMS Dev test mailbox - See Brian Stark"
while ($number -lt 2) {
    $uppercase = "ABCDEFGHKLMNOPRSTUVWXYZ".tochararray() 
    $lowercase = "abcdefghiklmnoprstuvwxyz".tochararray() 
    $Pnumber = "0123456789".tochararray() 
    $special = "$%&/()=?}{@#*+!".tochararray() 
    For ($i = 0; $i -le 20; $i++) {
        $password = ($uppercase | Get-Random -count 4) -join ''
        $password += ($lowercase | Get-Random -count 8) -join ''
        $password += ($Pnumber | Get-Random -count 4) -join ''
        $password += ($special | Get-Random -count 4) -join ''
        $passwordarray = $password.tochararray() 
        $scrambledpassword = ($passwordarray | Get-Random -Count 20) -join ''
    }
    $DisplayName = "SCTSDev" + $number
    $FirstName = "SCTSDev"
    $Lastname = "Mailbox" + $Number
    $SAM = $Name + $number
    $user = [pscustomobject]@{
        'SAM'      = $SAM
        'Password' = $scrambledpassword
    }
    $user | out-file c:\PS\SCTSDev.txt -Append
    New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Description $Description  -Path $OU -Enabled $True -ChangePasswordAtLogon $false  -AccountPassword (ConvertTo-SecureString $scrambledpassword -AsPlainText -force)  -CannotChangePassword:$true -PasswordNeverExpires $true -passThru
   <#$Groups = "ICMS Development Perth Admin Clerk","ICMS Development Perth Office Manager","ICMS Development Perth Clerk of Court","ICMS Development Perth Adoption Admin Clerk","ICMS Development Peterhead Admin Clerk","ICMS Development Peterhead Office Manager","ICMS Development Peterhead Clerk of Court","ICMS Development Peterhead Adoption Admin Clerk","ICMS Development Portree Admin Clerk","ICMS Development Portree Office Manager","ICMS Development Portree Clerk of Court","ICMS Development Portree Adoption Admin Clerk","ICMS Development Selkirk Admin Clerk","ICMS Development Selkirk Office Manager","ICMS Development Selkirk Clerk of Court","ICMS Development Selkirk Adoption Admin Clerk","ICMS Development Stirling Admin Clerk","ICMS Development Stirling Office Manager","ICMS Development Stirling Clerk of Court","ICMS Development Stirling Adoption Admin Clerk","ICMS Development Stornoway Admin Clerk","ICMS Development Stornoway Office Manager","ICMS Development Stornoway Clerk of Court","ICMS Development Stornoway Adoption Admin Clerk","ICMS Development Stranraer Admin Clerk","ICMS Development Stranraer Office Manager","ICMS Development Stranraer Clerk of Court","ICMS Development Stranraer Adoption Admin Clerk","ICMS Development Tain Admin Clerk","ICMS Development Tain Office Manager","ICMS Development Tain Clerk of Court","ICMS Development Tain Adoption Admin Clerk","ICMS Development Wick Admin Clerk","ICMS Development Wick Office Manager","ICMS Development Wick Clerk of Court","ICMS Development Wick Adoption Admin Clerk"
    ForEach ($Group in $Groups) {
    Write-Output $Group
            Add-ADGroupMember -Identity $Group -Members $SAM 
        }#>
        Set-ADUser -Identity $SAM -add @{"extensionattribute2" = "TST-PERS" }
    $number++
}
Write-Output "Accounts created"