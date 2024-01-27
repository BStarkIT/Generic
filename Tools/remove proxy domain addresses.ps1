#Author: Jos Lieben (OGD)
#Date: 13-09-2016
#Script home: www.lieben.nu
#Copyright: Leave this header intact, credit the author, otherwise free to use
#Purpose: Remove all *.onmicrosoft.com aliases from ALL active directory objects
#Requires –Version 3

$logPath = "c:\tools\results.log"
$readOnly = $False

#Check if elevated
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")){   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

function log{
    param (
        [Parameter(Mandatory=$true)][String]$text,
        [Parameter(Mandatory=$true)][String]$color,
        [Switch]$append
    )
    if($append){$text = "$text, $($Error[0].Exception) at $($Error[0].InvocationInfo.ScriptLineNumber) $($Error[0].ErrorDetails)"}
    Add-Content $logPath "$(Get-Date): $text"
    if($fout){
        Write-Host $text -ForegroundColor $color
    }else{
        Write-Host $text -ForegroundColor $color
    }
}

function cacheADObjects{
    try{
        [Array]$ADObjects = @(Get-ADObject -Filter * -Properties objectGuid,proxyAddresses,cn -ErrorAction Stop | Where-Object {$_.proxyAddresses})
    }catch{
        Throw "Failed to cache AD objects. $($Error[0]) $ADObjects"
    }
    return $ADObjects
}

log -text "$($env:USERNAME) on $($env:COMPUTERNAME) starting-----" -color "Green"
log -text "Log path: $logPath" -color "Green"

#region connect to local AD
try{
    Import-Module ActiveDirectory -Force -ErrorAction Stop
    log -text "Active Directory connection: OK" -color "Green"
}catch{
    log -text "failed to load Active Directory Module" -color "Red" -append
    Exit
}#endregion

log -text "Retrieving AD Objects, this may take a while..." -color "Green"
[array]$targetObjects = @(cacheADObjects | Where-Object {$_})
if($targetObjects.Count -gt 0){
    log -text "$($targetObjects.Count) objects found that have a proxyAddress configured" -color "Green"
}else{
    log -text "No objects retrieved from AD, exiting" -color "Red"
    Exit
}

foreach($object in $targetObjects){
    log -text "$($object.objectGuid) | $($object.cn) | current addresses: $($object.proxyAddresses -Join ",")" -color "Green"
    if($object.proxyAddresses -match ".Amazerealise.com"){
        $fixedProxyAddressesField = @()
        $fixedProxyAddressesField = $object.proxyAddresses -notmatch ".Amazerealise.com" 
        try{
            if(!$readOnly){
                $res = Set-ADObject -Identity $object.objectGuid -Replace @{proxyAddresses=$fixedProxyAddressesField}
            }
            log -text "$($object.objectGuid) | $($object.cn) | new addresses: $($fixedProxyAddressesField -Join ",")" -color "Green"
        }catch{
            log -text "$($object.objectGuid) | $($object.cn) | FAILED to set new addresses to: $($fixedProxyAddressesField -Join ",") $res" -color "Red" -append
        }
    }else{
        log -text "$($object.objectGuid) | $($object.cn) | no changes required" -color "Green"
    }
}