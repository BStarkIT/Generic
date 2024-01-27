$csvContent = Get-Content "\\scotcourts.local\home\P\BHoggard xx DELETE after DATE - 04 Apr 2023\UserMembershipBackup.csv"
Foreach ($group in $csvContent) {
Get-ADGroup $group

}
