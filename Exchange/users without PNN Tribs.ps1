$users = Import-Csv C:\PS\BBStaff.csv
$Shared = 'AllBlackberryUsers@scotcourts.gov.uk'
foreach ($User in $Users.Email) {
    Add-DistributionGroupMember -Identity $Shared -Member $user
}
