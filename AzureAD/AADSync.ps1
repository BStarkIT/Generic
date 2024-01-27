
<#
.SYNOPSIS
This PowerShell script is to force AD Connect sync to kick off from Active Directory to Office 365. Assuming you have AD Connect set up on one of your servers.

.NOTES
Script written by Brian Stark
Date: Pre 2022
Reviewed by:
Date:

change FQDN on line 4 to the server's FQDN name.

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>

$AADComputer = "FQDN"
$session = New-PSSession -ComputerName $AADComputer
Invoke-Command -Session $session -ScriptBlock {Import-Module -Name 'ADSync'}
Invoke-Command -Session $session -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
Remove-PSSession $session
