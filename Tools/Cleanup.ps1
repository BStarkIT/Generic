<#
.SYNOPSIS
Automated disk clean-up

.NOTES
Script written by Brian Stark
Date:
Reviewed by:
Date:

Stored in Project Repo:

.DESCRIPTION
written by BStark
#>
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace(0xA)
#1# Empty Recycle Bin #
write-Host "Emptying Recycle Bin." -ForegroundColor Cyan 
$objFolder.items() | ForEach-Object { remove-item $_.path -Recurse -Confirm:$false }
#2# Remove Temp
write-Host "Removing Temp" -ForegroundColor Green
Set-Location “$Env:TEMP”
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue
Set-Location “$env:windir\Temp\”
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue
Set-Location “$env:windir\Prefetch”
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue
Set-Location “$env:windir\Logs\CBS\”
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue
Set-Location “$env:windir\SoftwareDistribution\Download\”
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue
#3# Running Disk Clean up Tool 
write-Host "Finally now , Running Windows disk Clean up Tool" -ForegroundColor Cyan
cleanmgr /sagerun:1 | out-Null 
write-Host "CleanUp Task Has been Completed Successfully, Bye Bye(Tricknology)" -ForegroundColor Yellow 
Write-Output -InputObject "Press any key to continue..."
[void]$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")