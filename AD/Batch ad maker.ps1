<#
.SYNOPSIS
This PowerShell script is to batch make perftest accounts onprem, starting on a ste number saving each accounts password to a txt file.

.NOTES
Script written by Brian Stark
Date: pre 2022
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
$number = 6
$Name = "perftest_t"
$OU = "OU=AzureAD,OU=User Accounts (Testing),OU=SCTS,DC=scotcourts,DC=local"
$Description = "ICMS Performance Test Account"
while ($number -lt 9) {
    $uppercase = "ABCDEFGHKLMNOPRSTUVWXYZ".tochararray() 
    $lowercase = "abcdefghiklmnoprstuvwxyz".tochararray() 
    $Pnumber = "0123456789".tochararray() 
    $special = "$%&/()=?}{@#*+!".tochararray() 
    For ($i = 0; $i -le 10; $i++) {
        $password = ($uppercase | Get-Random -count 2) -join ''
        $password += ($lowercase | Get-Random -count 4) -join ''
        $password += ($Pnumber | Get-Random -count 2) -join ''
        $password += ($special | Get-Random -count 2) -join ''
        $passwordarray = $password.tochararray() 
        $scrambledpassword = ($passwordarray | Get-Random -Count 10) -join ''
    }
    $DisplayName = "Performance_Test" + $number
    $FirstName = "Performance"
    $Lastname = "Test" + $Number
    $SAM = $Name + $number
    $user = [pscustomobject]@{
        'SAM'      = $SAM
        'Password' = $scrambledpassword
    }
    $user | out-file c:\PS\Pref.txt -Append
    New-AdUser -Name "$DisplayName" -SamAccountName "$SAM" -GivenName "$FirstName" -Surname "$Lastname" -DisplayName "$DisplayName" -UserPrincipalName "$SAM@scotcourts.gov.uk" -Description $Description  -Path $OU -Enabled $True -ChangePasswordAtLogon $false  -AccountPassword (ConvertTo-SecureString $scrambledpassword -AsPlainText -force)  -CannotChangePassword:$true -PasswordNeverExpires $true -passThru
    $Groups = Get-ADgroup -filter { GroupCategory -eq "Security" -and Name -like "ICMS Development *" -and Name -NotLike "* Judiciary" }
    ForEach ($Group in $Groups) {
    Write-Output $Group
            Add-ADGroupMember -Identity $Group -Members $SAM 
        }
    Enable-MailBox -Identity $SAM@scotcourts.gov.uk
    Set-Mailbox -Identity $SAM@scotcourts.gov.uk -HiddenFromAddressListsEnabled $true
    $number++
}
Write-Output "Accounts created"