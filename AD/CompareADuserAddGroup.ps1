<#
.SYNOPSIS
This PowerShell script is to compare users and add to groups

.NOTES
Script written by Brian Stark
Date: pre 2022
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
param(
    [Parameter(Mandatory = $true)]
    [string] $SourceAcc,
    [Parameter(Mandatory = $true)]
    [string] $DestAcc,
    [string] $MatchGroup,
    [switch] $NoConfirm
)

# Retrieves the group membership for both accounts
$SourceMember = Get-AdUser -Filter { samaccountname -eq $SourceAcc } -Property memberof | Select-Object memberof
$DestMember = Get-AdUser -Filter { samaccountname -eq $DestAcc } -Property memberof | Select-Object memberof

# Checks if accounts have group membership, if no group membership is found for either account script will exit
if ($null -eq $SourceMember) { 'Source user not found'; return }
if ($null -eq $DestMember) { 'Destination user not found'; return }

# Uses -match to select a subset of groups to copy to the new user
if ($MatchGroup) {
    $SourceMember = $SourceMember | Where-Object { $_.memberof -match $MatchGroup }
}

# Checks for differences, if no differences are found script will prompt and exit
if (-not (Compare-Object $DestMember.memberof $SourceMember.memberof | Where-Object { $_.sideindicator -eq '=>' })) { write-host "No difference between $SourceAcc & $DestAcc groupmembership found. $DestAcc will not be added to any additional groups."; return }

# Routine that changes group membership and displays output to prompt
compare-object $DestMember.memberof $SourceMember.memberof | where-object { $_.sideindicator -eq '=>' } |
Select-Object -expand inputobject | ForEach-Object { write-host "$DestAcc will be added to:"([regex]::split($_, '^CN=|,OU=.+$'))[1] }

# If no confirmation parameter is set no confirmation is required, otherwise script will prompt for confirmation
if ($NoConfirm)	{
    compare-object $DestMember.memberof $SourceMember.memberof | where-object { $_.sideindicator -eq '=>' } | 
    Select-Object -expand inputobject | ForEach-Object { add-adgroupmember "$_" $DestAcc }
}

else {
    do {
        $UserInput = Read-Host "Are you sure you wish to add $DestAcc to these groups?`n[Y]es, [N]o or e[X]it"
        if (('Y', 'yes', 'n', 'no', 'X', 'exit') -notcontains $UserInput) {
            $UserInput = $null
            Write-Warning 'Please input correct value'
        }
        if (('X', 'exit', 'N', 'no') -contains $UserInput) {
            Write-Host 'No changes made, exiting...'
            exit
        }     
        if (('Y', 'yes') -contains $UserInput) {
            compare-object $DestMember.memberof $SourceMember.memberof | where-object { $_.sideindicator -eq '=>' } | 
            Select-Object -expand inputobject | ForEach-Object { add-adgroupmember "$_" $DestAcc }
        }
    }
    until ($null -ne $UserInput)
}