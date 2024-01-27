<#
.SYNOPSIS
This PowerShell script is to fetch computer information.

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.PARAMETER ComputerName
ComputerName may be your local host or remote server name or IP address.

.EXAMPLE
PS C:\> Get-PCinformation -ComputerName '0.0.0.0'
    Computer Name with Serial Number

.NOTES
Script written by Brian Stark of BStarkIT 

.DESCRIPTION
written by BStark

.LINK
Scripts can be found at:
https://github.com/BStarkIT 
#>
$ApplicationName = "VCS-Send-Mail-SCTS-PA"
$StartDate = Get-Date
$EndDate = $StartDate.AddYears(98)
$AzureApplicationObject = New-AzureADApplication -DisplayName $ApplicationName -AvailableToOtherTenants $false 
$AzureApplicationSecret = New-AzureADApplicationPasswordCredential -ObjectId $AzureApplicationObject.ObjectID -EndDate $EndDate
$AzureApplicationSecret.value
$AzureApplicationSecret.value | clip