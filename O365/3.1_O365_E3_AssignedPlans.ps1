Connect-AzureAD
(Get-AzureADUser -ObjectId slamb@scotcourts.gov.uk).assignedplans


$userUPN="gkhosa@scotcourts.gov.uk"
$AllLicenses=(Get-MsolUser -UserPrincipalName $userUPN).Licenses
$licArray = @()
for($i = 0; $i -lt $AllLicenses.Count; $i++)
{
$licArray += "License: " + $AllLicenses[$i].AccountSkuId
$licArray +=  $AllLicenses[$i].ServiceStatus
$licArray +=  ""
}
$licArray


(Get-MsolUser -UserPrincipalName gkhosa@scotcourts.gov.uk).Licenses[<LicenseIndexNumber>].ServiceStatus

(Get-MsolUser -UserPrincipalName gkhosa@scotcourts.gov.uk).Licenses.ServiceStatus

(Get-MsolUser -UserPrincipalName gkhosa@scotcourts.gov.uk).Licenses[0].ServiceStatus