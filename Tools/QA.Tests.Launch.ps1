Start-Process "powershell" -ArgumentList ("-ExecutionPolicy Bypass -noprofile "), "-command Invoke-pester \\scotcourts.local\data\CDiScripts\Scripts\W10DeploymentQA.Tests.ps1"
pause
