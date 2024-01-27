### BB Cleanup.ps1 By Brian Stark
### 06/02/2020 V1
### 
### Delete all Blackberry backup files older than 30 day
###
### Variables
$Path1 = "E:\Backup\bes"
$Path2 = "E:\Backup\Control" 
$Path3 = "E:\Backup\master"
$Path4 = "E:\Backup\model"
$Path5 = "E:\Backup\msdb"
$Daysback = "-15" # Days to go back
###
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($Daysback)
Get-ChildItem $Path1 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
Get-ChildItem $Path2 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
Get-ChildItem $Path3 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
Get-ChildItem $Path4 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
Get-ChildItem $Path5 | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item