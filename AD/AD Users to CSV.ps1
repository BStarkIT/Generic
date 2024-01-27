<#
.SYNOPSIS
This PowerShell script is to Export AD users to a .csv file

.NOTES
Script written by Brian Stark
Date: pre 2019
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>


$path = Split-Path -parent "C:\PS\*.*"
$LogDate = get-date -f yyyyMMddhhmm
$csvfile = $path + "\ADUsers_$logDate.csv"
Import-Module ActiveDirectory
##
## Change to match domain OU to be exported
##
$SearchBase = "OU=Realise Users ,DC=ad,DC=realise,DC=com"
##
$AllADUsers = Get-ADUser -searchbase $SearchBase -Properties * -Filter * |  
Select-Object @{Label = "First Name"; Expression = { $_.GivenName } },
@{Label = "Last Name"; Expression = { $_.Surname } },
@{Label = "Display Name"; Expression = { $_.DisplayName } },
@{Label = "Logon Name"; Expression = { $_.sAMAccountName } },
@{Label = "Full address"; Expression = { $_.StreetAddress } },
@{Label = "City"; Expression = { $_.City } },
@{Label = "State"; Expression = { $_.st } },
@{Label = "Post Code"; Expression = { $_.PostalCode } },
@{Label = "Country/Region"; Expression = { '' } },
@{Label = "Job Title"; Expression = { $_.Title } },
@{Label = "Company"; Expression = { $_.Company } },
@{Label = "Directorate"; Expression = { $_.Description } },
@{Label = "Department"; Expression = { $_.Department } },
@{Label = "Office"; Expression = { $_.OfficeName } },
@{Label = "Phone"; Expression = { $_.telephoneNumber } },
@{Label = "Email"; Expression = { $_.Mail } },
@{Label = "Manager"; Expression = { ForEach-Object { (Get-AdUser $_.Manager -server $ADServer -Properties DisplayName).DisplayName } } },
@{Label = "Account Status"; Expression = { if (($_.Enabled -eq 'TRUE')  ) { 'Enabled' } Else { 'Disabled' } } }, 
@{Label = "Last LogOn Date"; Expression = { $_.lastlogondate } } |
##
## Export CSV report
##
Export-Csv -Path $csvfile -NoTypeInformation
