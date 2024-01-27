$PC = read-Host 'Name of PC to be checked'
Write-Output "Checking the machine is online:"
Test-Connection -BufferSize 32 -Count 1 -ComputerName $PC -Quiet
Write-Output "Gathering installed software"
get-wmiobject Win32_Product -computername $PC | Format-Table IdentifyingNumber, Name, LocalPackage -AutoSize