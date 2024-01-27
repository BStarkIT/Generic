param
(
    [Parameter(Mandatory)]$Hrs,
    [Parameter(Mandatory)]$domain
)
Get-MessageTrackingLog -resultsize unlimited -Start (Get-Date).AddHours(-$Hrs) | Where-Object { $_.sender -like $domain }