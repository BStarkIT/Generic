<# Azure AD Sync
.SYNOPSIS
This PowerShell script is to Trigger an Azure AD Sync.

.NOTES
Script written by Brian Stark
Date: 2023
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
$Server = "SAUAZADC01"
$UserCredential = Get-Credential
Invoke-Command -ComputerName $Server -Credential $UserCredential -ScriptBlock { Start-ADSyncSyncCycle -PolicyType delta }