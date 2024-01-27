<#
.SYNOPSIS
This PowerShell script is to auto restart the Palo Alto UserID monitoring service

.NOTES
Script written by Brian Stark
Date: 17/03/2023
Reviewed by:
Date:

.DESCRIPTION
written by BStark
self eleavating 
Monitors evey 5 mins, restarts Palo Alto's service "User-ID Agent" if needed.
To be set as a schedule task to run as service account <X>
On servers SAUMSIT01 and GLWMSIT01 
Requested by Daniel Kelly.
#>
# Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
     $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
     Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
     Exit
    }
   }
   $ServiceName = 'User-ID Agent'
   while (1) { 
       $arrService = Get-Service -Name $ServiceName
       while ($arrService.Status -ne 'Running') {
           Start-Service $ServiceName
           write-host $arrService.status
           write-host 'Service starting'
           Start-Sleep -seconds 5
           $arrService.Refresh()
           if ($arrService.Status -eq 'Running') {
               Write-Host 'Service is now Running'
           }
       }
       Start-Sleep -seconds 3600
   }