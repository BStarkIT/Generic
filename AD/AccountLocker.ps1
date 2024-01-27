<# AD account locker
.SYNOPSIS
This PowerShell script is to lock a User account 

.NOTES
Script written by Brian Stark
Date:
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
#Requires -Version 3.0
#Requires -Modules ActiveDirectory, GroupPolicy
$Server = ""
$user = read-host "enter user id to lock out"
if ($LockoutBadCount = ((([xml](Get-GPOReport -Name "Default Domain Policy" -ReportType Xml)).GPO.Computer.ExtensionData.Extension.Account |
            Where-Object name -eq LockoutBadCount).SettingNumber)) {
    $Password = ConvertTo-SecureString 'NotMyPassword' -AsPlainText -Force
    Get-ADUser -identity $user -Properties SamAccountName, UserPrincipalName, LockedOut |
        ForEach-Object {
            for ($i = 1; $i -le $LockoutBadCount; $i++) { 
                Invoke-Command -ComputerName $Server {Get-Process
                } -Credential (New-Object System.Management.Automation.PSCredential ($($_.UserPrincipalName), $Password)) -ErrorAction SilentlyContinue            
            }
            Write-Output "$($_.SamAccountName) has been locked out: $((Get-ADUser -Identity $_.SamAccountName -Properties LockedOut).LockedOut)"
        }
}