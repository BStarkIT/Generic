Function RoundoffGB
{
 param($a)
 $b = $a -split 'B '| select -Last 1
 $c = $b.TrimEnd(" bytes)")
 $d = $c.TrimStart("(")
 $e = $d -replace (",","")
 $f = $e/1GB
 [math]::Round($f,2)
}
Function RoundoffMB
{
 param($a)
 $b = $a -split 'B '| select -Last 1
 $c = $b.TrimEnd(" bytes)")
 $d = $c.TrimStart("(")
 $e = $d -replace (",","")
 $f = $e/1MB
 [math]::Round($f,2)
 }
$mbxs=@()
$Table=@()
$count=1

$mbxs = Get-Mailbox -ResultSize unlimited
Foreach($mbx in $mbxs)
{
  $Aduser = $null
  $email = $mbx.PrimarySmtpAddress
  Write-progress -Activity "Working on the user :$Email " -status $("InProgress - "+$Count+"/"+@($mbxs).Count)-CurrentOperation $("") -PercentComplete (($Count/@($mbxs).Count)*100) -id 1
  If($mbx.ProhibitSendReceiveQuota -like "unlimited")
  {
    $Quota = "Unlimited"
  }
  Else
  {
    $Quota = RoundoffMB $(($mbx).ProhibitSendReceiveQuota)
  }
  If($mbx.ArchiveQuota -like "unlimited")
  {
    $AQuota = "Unlimited"
  }
  Else
  {
    $AQuota = RoundoffMB $(($mbx).ArchiveQuota)
  }
  $stat = Get-MailboxStatistics -identity $Email
  $AStat = Get-MailboxStatistics -identity $Email -Archive -ErrorAction SilentlyContinue
  $TotalItemSizeMB = Roundoffmb $($Stat.TotalItemSize)
  $TotalItemSizeGB = Roundoffgb $($Stat.TotalItemSize)
  $TotalDeletedItemSize = RoundoffGB $($stat.TotalDeletedItemSize)
  $ArchiveTotalItemSizeMB = Roundoffmb $($AStat.TotalItemSize)
  $ArchiveTotalItemSizeGB = Roundoffgb $($AStat.TotalItemSize)
 
  $ADUser = Get-ADUser -Filter 'mail -like $Email' -properties * -Server SAU-DC-02.scotcourts.local -Erroraction silentlycontinue
   
   If($ADuser -eq $null)
   {
       $ADuser = Get-ADuser $mbx.Name -properties * -Server SAU-DC-02.scotcourts.local -ErrorAction silentlycontinue
       If($ADuser -eq $null)
       {
            $ADuser = Get-ADuser $mbx.Alias -properties * -Server SAU-DC-02.scotcourts.local -ErrorAction silentlycontinue
       }
   }

  $upn = ($mbx.UserPrincipalName).Split('@')
  
  $proxy = $Aduser.proxyAddresses -join "~"
  $EmailAddresses = $mbx.EmailAddresses -join "~"
  $obj = New-Object PSCustomObject
  $obj | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mbx.PrimarySmtpAddress
  $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $mbx.DisplayName
  $obj | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mbx.UserPrincipalName
  $obj | Add-Member -MemberType NoteProperty -Name UPN_domain -Value $upn[1]
  $obj | Add-Member -MemberType NoteProperty -Name Name -Value $mbx.name
  $obj | Add-Member -MemberType NoteProperty -Name Alias -Value $mbx.Alias
  $obj | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $mbx.samaccountname
  $obj | Add-Member -MemberType NoteProperty -Name RecipientType -Value $mbx.RecipientType
  $obj | Add-Member -MemberType NoteProperty -Name RecipientTypedetails -Value $mbx.RecipientTypedetails
  $obj | Add-Member -MemberType NoteProperty -Name RecipientOU -Value $mbx.OrganizationalUnit
  $obj | Add-Member -MemberType NoteProperty -Name EmailAddresses -Value $EmailAddresses
  $obj | Add-Member -MemberType NoteProperty -Name Database -Value $mbx.Database
  $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $mbx.ServerName
  $obj | Add-Member -MemberType NoteProperty -Name UseDatabaseQuotaDefaults -Value $mbx.UseDatabaseQuotaDefaults
  $obj | Add-Member -MemberType NoteProperty -Name ProhibitSendReceiveQuota-MB -Value $quota
  $obj | Add-Member -MemberType NoteProperty -Name TotalItemSizeMB -Value $TotalItemSizeMB
  $obj | Add-Member -MemberType NoteProperty -Name TotalItemSizeGB -Value $TotalItemSizeGB
  $obj | Add-Member -MemberType NoteProperty -Name TotalDeletedItemSize -Value $TotalDeletedItemSize
  $obj | Add-Member -MemberType NoteProperty -Name ItemCount -Value $stat.ItemCount
  $obj | Add-Member -MemberType NoteProperty -Name DeletedItemCount -Value $stat.DeletedItemCount
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveName -Value $mbx.Archivename
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveStatus -Value $mbx.ArchiveStatus
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveState -Value $mbx.ArchiveState
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveQuota-MB -Value $AQuota
  $obj | Add-Member -MemberType NoteProperty -Name LastLogonTime -Value $stat.LastLogonTime
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveTotalItemSizeMB -Value $ArchiveTotalItemSizeMB
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveTotalItemSizeGB -Value $ArchiveTotalItemSizeGB
  $obj | Add-Member -MemberType NoteProperty -Name ArchiveTotalItemCount -Value $Astat.ItemCount
  $obj | Add-Member -MemberType NoteProperty -Name legacyexchangedn -Value $mbx.legacyexchangedn
  $obj | Add-Member -MemberType NoteProperty -Name distinguishedname -Value $mbx.distinguishedname
  $obj | Add-Member -MemberType NoteProperty -Name Employeeid -Value $aduser.employeeid
  $obj | Add-Member -MemberType NoteProperty -Name Employeetype -Value $aduser.employeetype
  $obj | Add-Member -MemberType NoteProperty -Name customattribute1 -Value $mbx.customattribute1
  $obj | Add-Member -MemberType NoteProperty -Name customattribute2 -Value $mbx.customattribute2
  $obj | Add-Member -MemberType NoteProperty -Name customattribute3 -Value $mbx.customattribute3
  $obj | Add-Member -MemberType NoteProperty -Name customattribute14 -Value $mbx.customattribute14
  $obj | Add-Member -MemberType NoteProperty -Name customattribute15 -Value $mbx.customattribute15
  $obj | Add-Member -MemberType NoteProperty -Name ADEnabled -Value $aduser.Enabled
  $obj | Add-Member -MemberType NoteProperty -Name msExchRecipientDisplayType -Value $aduser.msExchRecipientDisplayType
  $obj | Add-Member -MemberType NoteProperty -Name msExchRecipientTypeDetails -Value $aduser.msExchRecipientTypeDetails
  $obj | Add-Member -MemberType NoteProperty -Name msExchRemoteRecipientType -Value $aduser.msExchRemoteRecipientType
  $obj | Add-Member -MemberType NoteProperty -Name HiddenFromAddressListsEnabled -Value $mbx.HiddenFromAddressListsEnabled
  $obj | Add-Member -MemberType NoteProperty -Name emailaddresspolicyenabled -Value $mbx.emailaddresspolicyenabled
  $obj | Add-Member -MemberType NoteProperty -Name msExchHideFromAddressLists -Value $aduser.msExchHideFromAddressLists 
  $obj | Add-Member -MemberType NoteProperty -Name TargetAddress -Value $aduser.TargetAddress
  $obj | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $mbx.ForwardingAddress -Force
  $obj | Add-Member -MemberType NoteProperty -Name ForwardingSmtpAddress -Value $mbx.ForwardingSmtpAddress -Force
  If($Aduser.TargetAddress -eq $null)
  {
    $obj | Add-Member -MemberType NoteProperty -Name O365orOnprem -Value "OnPrem"
  }
  ElseIf($Aduser.TargetAddress -like "*scotcourtsgovuk.mail.onmicrosoft.com")
  {
  $obj | Add-Member -MemberType NoteProperty -Name O365orOnprem -Value "O365"
  }
  $obj | Add-Member -MemberType NoteProperty -Name proxyAddresses -Value $proxy
  $Table +=$obj
  $count++
  }

Write-Host "`n`n"
Write-Host "Exporting the results to CSV file" -ForegroundColor Green
$Dat = "EX2016_onprem" + "_" + (Get-Date -Format "ddMMMyyyhhmmss") + ".csv"
$table | export-csv C:\scripts\Exchange\reports\$Dat -NoTypeInformation -Force
Write-Host " Results has been exported to C:\scripts\Exchange\reports\$Dat" -ForegroundColor green
