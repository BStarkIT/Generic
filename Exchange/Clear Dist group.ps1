$list = Get-DistributionGroupMember -Identity $DistributionGroup
$list | % {
   Remove-DistributionGroupMember -Identity $DistributionGroup -Member $_.Name -Confirm:$false
   } 