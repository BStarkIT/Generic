$number = (Get-PhysicalDisk | Select-Object -Property Uniqueid).Uniqueid
write-host $number
