<#
.SYNOPSIS
This PowerShell script is to change the setting of accounts which names begin with "HMP" & are located within the Resource mailbox OU.

.NOTES
Written by Vicky Pasieka    -   07/12/2022
Proofed by Brian Stark      -   07/12/2022
Approved on  CAB            -

.DESCRIPTION
written by VPasieka

#>
$collection = Get-ADUser -filter 'name -Like "HMP*" ' -searchbase 'OU=ResourceMailboxes,DC=scotcourts,DC=local' -Properties sAMAccountName| Select-Object sAMAccountName | Select-Object -ExpandProperty sAMAccountName
foreach ($HMP in $collection) {
    Set-ADUser -Identity $HMP -Clear msExchHideFromAddressLists
    Write-Output "Account $HMP changed"
}
