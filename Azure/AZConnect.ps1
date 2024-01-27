$SessionAzure = New-PSSession -ComputerName SAUAZADC01 -Credential $UserCredential
Enter-PSSession $SessionAzure
Start-Sleep -seconds 5
Start-ADSyncSyncCycle -PolicyType Delta