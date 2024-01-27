$Hrs = read-Host 'Number of hrs to go back'
$domain = read-Host 'domain to search for'
Get-MessageTrackingLog -resultsize unlimited -Start (Get-Date).AddDays(-$Hrs) | Where-Object { $_.Recipients -like "*$domain*" }