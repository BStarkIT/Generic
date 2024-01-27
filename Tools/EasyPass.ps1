$Description = Import-Csv .\Pass.csv | Get-Random | Select-Object -ExpandProperty Description
$Object = Import-Csv .\Pass.csv | Get-Random | Select-Object -ExpandProperty Object
$Complex1 = Import-Csv .\Pass.csv | Get-Random | Select-Object -ExpandProperty Complex
$Complex2 = Import-Csv .\Pass.csv | Get-Random | Select-Object -ExpandProperty Complex
Write-host "Password is : $Description$Complex1$Object$Complex2"
