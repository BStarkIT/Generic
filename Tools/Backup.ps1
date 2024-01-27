Write-Host "Copying Document files"
robocopy "$env:UserProfile\Documents" \\control\paperwork\Docs /E /Z
Write-Host "Copying Picture files"
robocopy "$env:UserProfile\Pictures" \\control\paperwork\Pics /E /Z
Write-Host "Copying Download files"
robocopy "$env:UserProfile\Downloads" \\control\paperwork\Downloads /E /Z
Write-Host "Copying Script files"
robocopy "$env:UserProfile\Git" \\control\paperwork\Git /E /Z
Write-Host "Updating Document files"
robocopy \\control\paperwork\Docs "$env:UserProfile\Documents" /E /Z
Write-Host "Updating Picture files"
robocopy \\control\paperwork\Pics "$env:UserProfile\Pictures" /E /Z
Write-Host "Updating Download files"
robocopy \\control\paperwork\Downloads "$env:UserProfile\Downloads" /E /Z
Write-Host "Updating Script files"
robocopy \\control\paperwork\Git "$env:UserProfile\Git" /E /Z
