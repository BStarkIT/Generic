<#
.SYNOPSIS
This PowerShell script is to pull the disabled user list from Azure AD, check they have been disabled via the leavers script, if not it will give a list to check

.NOTES
Script written by Brian Stark
Date: 05/05/2023
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
Connect-MgGraph
$EndDate = (Get-Date).ToString('dd MMM yyyy')
$newsheet = @()
$Outputsheet = @()
$Contractor = @()
$users = Get-MgGroupMember -GroupId 2ffecedc-fb37-4288-91fb-48264ee22cdf 
[array]$allmemebers = $users.AdditionalProperties
foreach ($Staff in $allmemebers) {
    $upn = $Staff.userPrincipalName
    $Sam = $upn.Split("@")[0]
    $details = Get-ADUser $Sam -Properties * | Select-Object Description, CanonicalName, CN
    $description = $details.Description
    $location2 = $details.CanonicalName
    $cn = $details.CN
    $Location1 = $location2 -replace ($cn, '')
    $location = $location1 -replace ('scotcourts.local/', '')
    if ($description.contains("xx DELETE after DATE --")) {
        $Marked = $true
    }
    else {
        $Marked = $false
    }
    if ($location.contains("Z-Disabled_Leavers")) {
        $path = $true
    }
    else {
        $path = $false
    }
    $disabledinfo = [PSCustomObject]@{
        Sam = $sam; Path = $path; Marked = $Marked; FullPath = $location; Description = $description
    }
    $newsheet += $disabledinfo 
}
foreach ($NRow in $newsheet) {
    if (($Nrow.Path -eq $false) -or ($Nrow.Marked -eq $false)) {
        $disablederror = [PSCustomObject]@{
            Sam = $Nrow.sam; Object = $Nrow.FullPath; Description = $Nrow.description
        }
        $exp = Get-ADUser $Nrow.Sam -Properties * | select-object -ExpandProperty AccountExpirationDate
        if ($null -eq $exp) {
            $Outputsheet += $disablederror
        }
        elseif ($exp -lt $EndDate) {
            $disabledCont = [PSCustomObject]@{
                Contractor = $Nrow.sam; OU = $Nrow.FullPath; Description = $Nrow.description
            }
            $Contractor += $disabledCont
        }
    }
}
Write-Output "the following account are disabled"
Write-Output $newsheet | Format-Table
Write-Output "the following accounts are contractors"
Write-Output $Contractor | Format-Table
Write-Output "the following accounts need checked"
Write-Output $Outputsheet | Format-Table


