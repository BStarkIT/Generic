#emailaddress as the csv column header
$Inputdata=(Import-csv c:\scripts\O365\migration_precomp_check\daily_migcomplete_check_batch10_10062022.csv).EmailAddress
$count=1
$Table=@()
Foreach($Email in $InputData)
{
  $stat = $null
  Write-progress -Activity "Working on the user :$Email " -status $("InProgress - "+$Count+"/"+@($Inputdata).Count)-CurrentOperation $("") -PercentComplete (($Count/@($inputdata).Count)*100) -id 1
  Try
  {
    $Stat=Get-Migrationuserstatistics -identity $Email -Erroraction Stop
  }
  Catch
  {
    Write-Host "$_" -ForegroundColor Red
  }
  If($Stat -ne $null)
  {
      $obj = New-Object PSCustomObject
      $obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $Email
      $obj | Add-Member -MemberType NoteProperty -Name Status -Value $Stat.Status
      $obj | Add-Member -MemberType NoteProperty -Name StatusSummary -Value $Stat.StatusSummary
      $obj | Add-Member -MemberType NoteProperty -Name PercentageComplete -Value $Stat.PercentageComplete
      $obj | Add-Member -MemberType NoteProperty -Name RecipientType -Value $Stat.RecipientType
      $obj | Add-Member -MemberType NoteProperty -Name BatchId -Value $Stat.BatchId
      $obj | Add-Member -MemberType NoteProperty -Name RecipientTypeDetails -Value $Stat.RecipientTypeDetails
      $obj | Add-Member -MemberType NoteProperty -Name Itemsmigrated -Value $Stat.SyncedItemCount
      $obj | Add-Member -MemberType NoteProperty -Name SkippedItemCount -Value $Stat.SkippedItemCount
      If($stat.PercentageComplete -eq '95' -and $stat.Status -like "Synced")
      {
        $obj | Add-Member -MemberType NoteProperty -Name ReadytoComplete -Value "True"
      }
      Else
      {
        $obj | Add-Member -MemberType NoteProperty -Name ReadytoComplete -Value "False"
      }
      $obj | Add-Member -MemberType NoteProperty -Name Error -Value $Stat.Error
      $obj | Add-Member -MemberType NoteProperty -Name Errorsummary -Value $Stat.Errorsummary
      $Table+=$obj
  }
  Else
  {
    Write-Host "$Email : Issue while pulling the reports" -ForegroundColor Red
  }
  $count++
}
If($Table -ne $null)
{
   Write-Host "`n"
   Write-Host "Exporting the results to CSV file" -ForegroundColor Green
   $Dat = "Migration_data_batch10" + "_" + (Get-Date -Format "ddMMMyyyhhmmss") + ".csv"
   $Table| Export-csv C:\scripts\O365\migration_precomp_report\$Dat -NoTypeInformation -Force
   Write-Host " Results has been exported to c:\scripts\O365\migration_precomp_report\$Dat" -ForegroundColor green
}
