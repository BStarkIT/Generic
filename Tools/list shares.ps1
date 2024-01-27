$Syst = read-Host 'Name of System to be checked'
Invoke-Command -ComputerName $Syst -ScriptBlock {Get-SmbShare}